"""Build editable DOCX, linked HTML and A4 PDFs from the reviewed JSON sources.

Requires Python python-docx, playwright, PyMuPDF, system Graphviz and Chromium.
Run from any directory: python tools/build_books.py [12 13 14]
"""
import argparse
import asyncio
import html
import json
import re
import subprocess
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from playwright.async_api import async_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'output'
TITLES = {12: 'ژنتیک پروکاریوت‌ها', 13: 'ویروس‌شناسی پیشرفته',
          14: 'گزیده آزمون‌های سنوات با پاسخ و تحلیل'}
NUMBERS = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')
LATIN_RE = re.compile(r"[+−\-]?[A-Za-z0-9\u0370-\u03ff\u06f0-\u06f9][A-Za-z0-9\u0370-\u03ff\u06f0-\u06f9\s._,:;/\\+=<>′’'−–\-→←±×()%\[\]{}^٫٬]*")
STYLE = '''
@font-face{font-family:Yasi;src:url('../assets/fonts/BNazanin.ttf')}
@page{size:A4;margin:15mm 15mm 17mm 25mm}
@page:left{margin-left:15mm;margin-right:25mm}
*{box-sizing:border-box}html{direction:rtl}body{font-family:Yasi,'Noto Naskh Arabic',serif;
font-size:14pt;line-height:1.48;color:#18202b;margin:0}h1,h2,h3,h4{line-height:1.35;
break-after:avoid;color:#173b5a}h1{font-size:25pt}h2{font-size:20pt;margin:16pt 0 10pt}
h3{font-size:17pt;margin:14pt 0 8pt}h4{font-size:15pt;margin:9pt 0 4pt}
p{margin:5pt 0 7pt;orphans:3;widows:3;text-align:justify}bdi,.latin{font-family:'DejaVu Sans',sans-serif;
font-size:.78em;direction:ltr;unicode-bidi:isolate}a{color:#17628b;text-decoration:none}
.chapter,.major{break-before:page}.meta,.citation{font-size:11.5pt;color:#526171}
.note{background:#f2f8fb;border-right:3px solid #208a9b;padding:9pt 12pt;margin:10pt 0}
.example{background:#f7f9fc;border:1px solid #d9e3e9;padding:9pt 12pt;margin:9pt 0;
break-inside:avoid}.question{break-inside:avoid;margin:10pt 0 12pt}.question p{margin-bottom:4pt}
.options{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:3pt 16pt;
padding:2pt 8pt}.options.long{grid-template-columns:1fr}.option{text-align:right}
.answer{break-inside:avoid;margin:7pt 0 9pt}.answer p{text-align:justify}
table{width:100%;border-collapse:collapse;margin:9pt 0;font-size:12.5pt}
thead{display:table-header-group}th{background:#e7f3f8;color:#173b5a}
td,th{border:1px solid #cddce5;padding:5pt 7pt;text-align:right;vertical-align:top}
tr{break-inside:avoid}figure{margin:12pt 0;break-inside:avoid;text-align:center}
figure img{width:100%;max-height:80mm;object-fit:contain}figcaption{font-size:11.5pt;
line-height:1.4;margin-top:5pt;color:#3d4d5a}.source{break-inside:avoid;margin:9pt 0;
overflow-wrap:anywhere;font-size:11.5pt}.toc p{margin:6pt 0}.keygrid{font-size:12pt}
.cover{width:210mm;height:297mm;position:relative;overflow:hidden;margin:0;
background:#14b4bd url('../assets/cover-art.png') center/cover no-repeat;
display:flex;flex-direction:column;align-items:center;text-align:center;padding:22mm 17mm 15mm}
.cover *{text-align:center}.cover .pill{border:2px solid #e3be5c;border-radius:22pt;
background:#431276;color:white;padding:6pt 20pt;font-size:24pt;line-height:1.35}
.cover h1{color:#39115c;font-size:43pt;line-height:1.32;margin:22mm 0 8mm;
text-shadow:0 1px 8px #fff;background:#fff9e8d9;padding:5mm;border-radius:8mm;width:100%}
.cover .subtitle{font-size:21pt;color:#431276;background:#ffe76be8;border:1px solid #d7a629;
padding:5pt 14pt;border-radius:20pt}.cover .author{font-size:23pt;background:#fffcebdc;
color:#431276;border:1px solid #d7b64d;border-radius:20pt;padding:4pt 16pt;margin-top:13mm}
.cover .badge{margin-top:auto;color:white;background:#c52e46;border:2px solid #e7c44f;
font-size:17pt;border-radius:22pt;padding:4pt 20pt}.cover.back{padding-top:30mm}
.back .backpanel{background:#fff9e9ed;border:3px double #daa938;border-radius:10mm;
padding:8mm;width:100%;font-size:20pt;line-height:1.7;color:#431276}
.backpanel h2{font-size:24pt;color:#431276;margin-top:0}.backpanel p{text-align:right}
@media screen{body{max-width:180mm;margin:20px auto;background:white}html{background:#edf2f5}}
'''


def fa(n):
    return str(n).translate(NUMBERS)


def esc(value):
    # Isolate Latin runs so formula/identifier order is stable within Persian text.
    value = str(value)
    output = []; end = 0
    for match in LATIN_RE.finditer(value):
        token = match.group().rstrip()
        output.append(html.escape(value[end:match.start()]))
        output.append('<bdi>' + html.escape(token) + '</bdi>')
        end = match.start() + len(token)
    output.append(html.escape(value[end:]))
    return ''.join(output)


def para(text, cls=''):
    return f'<p class="{cls}">{esc(text)}</p>'


def refs_html(ids):
    return '؛ '.join(f'<a href="#ref-{html.escape(i)}"><bdi>[{html.escape(i)}]</bdi></a>' for i in ids)


def question_html(q, number=None):
    label = fa(number) if number is not None else q['id']
    long = max(map(len, q['options'])) > 65 or sum(map(len, q['options'])) > 170
    options = ''.join(f'<div class="option">{fa(i+1)}) {esc(t)}</div>' for i, t in enumerate(q['options']))
    return (f'<div class="question" id="{html.escape(q["id"])}"><p><strong>{esc(label)} — {esc(q["stem"])}</strong></p>'
            f'<div class="options {"long" if long else ""}">{options}</div></div>')


def answer_html(q, number=None):
    label = fa(number) if number is not None else q['id']
    return (f'<div class="answer"><p><strong>{esc(label)} — گزینه {fa(q["answer"]+1)}.</strong> '
            f'{esc(q["explanation"])}</p><div class="citation">{refs_html(q["sources"])}</div></div>')


def table_html(headers, rows):
    return '<table><thead><tr>' + ''.join(f'<th>{esc(t)}</th>' for t in headers) + '</tr></thead><tbody>' + ''.join(
        '<tr>' + ''.join(f'<td>{esc(t)}</td>' for t in row) + '</tr>' for row in rows) + '</tbody></table>'


def make_diagram(volume, step):
    """Render authored semantic diagrams with Graphviz/Pango Arabic shaping."""
    directory = ROOT / 'figures' / str(volume)
    directory.mkdir(parents=True, exist_ok=True)
    base = directory / f'step-{step["number"]:02d}'
    d = step['diagram']
    lines = ['digraph G {', 'graph [rankdir=LR, bgcolor="white", pad="0.2", nodesep="0.35", ranksep="0.45", size="7,2.7", ratio="compress"];',
             'node [shape=box, style="rounded,filled", fillcolor="#eaf5fa", color="#247c91", fontname="Noto Naskh Arabic", fontsize=16, margin="0.16,0.12"];',
             'edge [color="#7454a5", arrowsize=0.8, fontname="Noto Naskh Arabic", fontsize=12];']
    for i, n in enumerate(d['nodes']):
        label = n['fa'] + '\n' + n['en']
        lines.append(f'n{i} [label={json.dumps(label, ensure_ascii=False)}];')
    for e in d['edges']:
        label = f' [label={json.dumps(e[2], ensure_ascii=False)}]' if len(e) == 3 and e[2] else ''
        lines.append(f'n{e[0]} -> n{e[1]}{label};')
    lines.append('}')
    base.with_suffix('.dot').write_text('\n'.join(lines))
    for fmt in ('svg', 'png'):
        subprocess.run(['dot', '-T' + fmt, '-Gdpi=180', str(base.with_suffix('.dot')),
                        '-o', str(base.with_suffix('.' + fmt))], check=True, capture_output=True)
    return base


def load_book(volume):
    if volume == 14:
        return [json.loads(p.read_text()) for p in sorted((ROOT / 'volumes/14').glob('questions-*.json'))]
    parts = [json.loads((ROOT / f'volumes/{volume}/part-{p}.json').read_text()) for p in ('a', 'b')]
    return {key: sum((d[key] for d in parts), []) for key in ('chapters', 'review', 'exams', 'glossary', 'sources')}


def sources_html(sources):
    rows = []
    for s in sources:
        rows.append(f'<div class="source" id="ref-{html.escape(s["id"])}"><strong><bdi>{html.escape(s["id"])}</bdi> — '
                    f'{esc(s["title"])}</strong>{para(s["section"])}'
                    f'<a href="{html.escape(s["url"], quote=True)}"><bdi>{html.escape(s["url"])}</bdi></a>'
                    f'{para("دامنه مشاهده: " + s["coverage"] + "؛ تاریخ دسترسی: " + s.get("accessed", "2026-11-20"))}</div>')
    return '<h2 class="major" id="sources">منابع و محل استناد</h2>' + ''.join(rows)


def teaching_html(volume, book):
    h = [f'<h1>جلد {fa(volume)} — {TITLES[volume]}</h1>',
         para('آمادگی آزمون دکتری ۱۴۰۶ | زیست‌شناسی سلولی و مولکولی و میکروبیولوژی', 'meta'),
         '<h2>پیشگفتار و راهنمای مطالعه</h2>',
         para('این جلد بر استدلال سازوکاری و تفسیر داده تمرکز دارد. مبانی در جلدهای پیشین آمده‌اند؛ اینجا هر گام با درسنامه، مثال حل‌شده، شکل تألیفی و سه تست همراه است. ابتدا سؤال را مستقل حل کنید و سپس پاسخ و منبع را بخوانید. مثال‌های عددی و نمودارهای این جلد مدل آموزشی‌اند و داده تجربی منتشرشده نامیده نمی‌شوند.'),
         para('ساختار: ده فصل و چهل گام، پنجاه سؤال مرور مستقل و چهار آزمون مستقل ۲۵ سؤالی. پاسخ هر گام بلافاصله پس از سؤال‌های آن است؛ پاسخ بانک مرور پس از بانک و پاسخ هر چهار آزمون پس از آزمون چهارم قرار دارد. شکل‌ها شماتیک و دوزبانه‌اند؛ اندازه و فاصله اجزا مقیاس واقعی نیست.'),
         '<div class="note">نسخه آزمایشی غیرقابل انتشار. مواد و ضرایب رسمی ۱۴۰۶ هنوز با صفحه ۳۳ دفترچه سازمان سنجش تطبیق قطعی نشده‌اند؛ متن این جلد پیش‌بینی سهم سؤال‌ها یا تأیید رسمی آن سازمان نیست.</div>',
         '<h2 id="contents">فهرست پیونددار</h2><div class="toc">']
    for c in book['chapters']:
        h.append(f'<p><a href="#chapter-{c["number"]}">فصل {fa(c["number"])} — {esc(c["title"])}</a></p>')
    h += ['<p><a href="#review">بانک پنجاه سؤال مرور</a></p><p><a href="#exams">چهار آزمون جامع</a></p>',
          '<p><a href="#exam-answers">کلید و پاسخ تشریحی چهار آزمون</a></p><p><a href="#glossary">واژه‌نامه</a></p><p><a href="#sources">منابع</a></p></div>']
    for c in book['chapters']:
        h.append(f'<h2 class="chapter" id="chapter-{c["number"]}">فصل {fa(c["number"])} — {esc(c["title"])}</h2>')
        for s in c['steps']:
            h.append(f'<h3 id="V{volume}-S{s["number"]:02d}">گام {fa(s["number"])} — {esc(s["title"])}</h3>')
            h.extend(para(p) for p in s['lesson'])
            ex = s['example']
            h.append('<div class="example"><strong>مثال حل‌شده</strong>' + para(ex['problem']) + para(ex['solution']) + '</div>')
            base = make_diagram(volume, s)
            uri = '../' + str(base.with_suffix('.svg').relative_to(ROOT))
            h.append(f'<figure><img src="{uri}" alt="{html.escape(s["diagram"]["title"], quote=True)}">'
                     f'<figcaption>شکل {fa(s["number"])} — {esc(s["diagram"]["title"])} (شماتیک تألیفی)</figcaption></figure>')
            h.append(f'<p class="citation">منابع این گام: {refs_html(s["sources"])}</p><h4>تست‌های همین گام — تألیفی</h4>')
            h.extend(question_html(q, i+1) for i, q in enumerate(s['questions']))
            h.append('<h4>پاسخ تشریحی همین گام</h4>')
            h.extend(answer_html(q, i+1) for i, q in enumerate(s['questions']))
        h.append('<h3>خلاصه و نکات مرور فصل</h3>')
        h.extend(para('• ' + p) for p in c['summary'])
        h.append(table_html(c['table']['headers'], c['table']['rows']))
    h.append('<h2 class="major" id="review">بانک پنجاه سؤال مرور مستقل — تألیفی</h2>')
    h.extend(question_html(q, i+1) for i, q in enumerate(book['review']))
    h.append('<h2 class="major">پاسخ تشریحی بانک مرور</h2>')
    h.extend(answer_html(q, i+1) for i, q in enumerate(book['review']))
    h.append('<h2 class="major" id="exams">چهار آزمون جامع مستقل — تألیفی</h2>')
    h.append(para('برای خودسنجی، هر آزمون را پیش از دیدن کلید حل کنید. زمان تمرین پیشنهادی ۳۵ دقیقه است و زمان رسمی آزمون نیست. نمره تمرینی با فرض یک نمره صحیح و یک‌سوم نمره منفی برای هر پاسخ غلط محاسبه می‌شود؛ بی‌پاسخ صفر است. این فرض را با مقررات آزمون مقصد تطبیق دهید.'))
    for e in book['exams']:
        h.append(f'<h2 class="major" id="exam-{e["number"]}">آزمون {fa(e["number"])} — {esc(e["title"])}</h2>')
        h.extend(question_html(q, i+1) for i, q in enumerate(e['questions']))
    h.append('<h2 class="major" id="exam-answers">کلید و پاسخ تشریحی چهار آزمون</h2>')
    for e in book['exams']:
        h.append(f'<h3>پاسخ آزمون {fa(e["number"])}</h3>')
        h.append(table_html(['سؤال', 'گزینه', 'سؤال', 'گزینه', 'سؤال', 'گزینه', 'سؤال', 'گزینه', 'سؤال', 'گزینه'],
            [[x for j in range(k, k+5) for x in (fa(j+1), fa(e['questions'][j]['answer']+1))] for k in range(0, 25, 5)]))
        h.extend(answer_html(q, i+1) for i, q in enumerate(e['questions']))
    h.append('<h2 class="major" id="glossary">واژه‌نامه فارسی–انگلیسی</h2>')
    h.append(table_html(['فارسی', 'English', 'تعریف آموزشی'], [[g['fa'], g['en'], g['definition']] for g in book['glossary']]))
    h.append(sources_html(book['sources']))
    return ''.join(h)


def exam_html(book):
    h = ['<h1>جلد ۱۴ — گزیده آزمون‌های سنوات</h1>', '<h2>دامنه و روش استفاده</h2>',
         para('این جلد شامل ۶۰ سؤال واقعی منتخب از دفترچه‌های ۱۴۰۳، ۱۴۰۴ و ۱۴۰۵ مجموعه ۲۲۲۶ است؛ هر سال ۲۰ سؤال از بخش مشترک، سلولی‌مولکولی پیشرفته و بخش میکروبی انتخاب شده است. این گزیده بانک کامل سال‌های ۱۳۹۵ تا ۱۴۰۵ نیست و فراوانی مباحث آن تخمین سهم آزمون آینده محسوب نمی‌شود. سؤال‌های رشته‌های دیگر وارد این انتخاب نشده‌اند.'),
         para('در هر سال ابتدا صورت و گزینه‌ها با ترتیب دفترچه برای تمرین آمده و سپس پاسخ تشریحی مستقل ارائه می‌شود. شماره اصلی سؤال حفظ شده است. فایل‌های تصویر در منبع قابل ویرایش و مخزن، شاهد رونویسی‌اند. فاصله‌گذاری و شکل حروف یکدست شده‌اند؛ بازنویسی مفهومی سؤال انجام نشده است.'),
         '<div class="note">پاسخ آموزشی با کلید رسمی تفاوت دارد. عنوان «رونویسی‌شده، تأییدنشده» یعنی تصویر کلید دیده شده اما انتساب آن به کد دفترچه و اولیه یا نهایی بودن آن قطعی نیست. «نامعلوم» یعنی کلید معتبر برای آن رکورد در دست نیست. اختلاف کلید و پاسخ در همان رکورد ثبت شده است.</div>',
         '<h2>فهرست</h2>']
    for d in book:
        h.append(f'<p><a href="#year-{d["year"]}">سال {fa(d["year"])} — {fa(len(d["questions"]))} سؤال</a></p>')
    h.append('<p><a href="#topic-index">نمایه موضوعی و محل مرور</a></p><p><a href="#sources">منابع علمی</a></p>')
    all_sources = []
    for d in book:
        y = d['year']; b = d['booklet']
        h.append(f'<h2 class="major" id="year-{y}">سال {fa(y)} — سؤال‌های منتخب</h2>')
        h.append(para(d['selection_note']))
        h.append(para(f'دفترچه: {b["title"]}؛ کد: {b["code"]}؛ تعداد صفحات PDF: {fa(b["pages"])}.'))
        h.append(f'<p class="meta">منبع دفترچه: <a href="{html.escape(b["source_url"], quote=True)}"><bdi>{html.escape(b["source_url"])}</bdi></a></p>')
        for q in d['questions']:
            h.append(question_html(q, q['number']))
            if q.get('figure_image'):
                h.append(f'<figure><img src="../{html.escape(q["figure_image"], quote=True)}" alt="شکل اصلی سؤال {fa(q["number"])}"><figcaption>شکل اصلی سؤال {fa(q["number"])}، سال {fa(y)}؛ برگرفته از همان دفترچه.</figcaption></figure>')
            elif q.get('requires_image'):
                raise ValueError(f'{q["id"]}: figure required but not supplied')
            h.append(para(f'صفحه PDF: {fa(q["pdf_page"])}؛ صفحه چاپی: {q.get("printed_page", "نامعلوم")}؛ مبحث: {q["topic"]}', 'meta'))
        h.append(f'<h2 class="major">پاسخ و تحلیل سال {fa(y)}</h2>')
        for q in d['questions']:
            h.append(answer_html(q, q['number']))
            key = q['key']; val = 'نامعلوم' if key.get('value') is None else f'گزینه {fa(key["value"]+1)}'
            status = {'unknown': 'نامعلوم', 'transcribed-unverified': 'رونویسی‌شده، تأیید رسمی/نهایی نامعلوم',
                      'verified': 'تطبیق‌شده با سند مشخص'}[key['status']]
            h.append(para(f'کلید منبع: {val}؛ وضعیت: {status}. {key.get("note", "")} محل مرور: {q["review_in"]}', 'meta'))
        all_sources.extend(d['sources'])
    h.append('<h2 class="major" id="topic-index">نمایه موضوعی و محل مرور</h2>')
    h.append(table_html(['سال/سؤال', 'مبحث', 'مسیر', 'محل مرور'], [[f'{fa(d["year"])}/{fa(q["number"])}',
        q['topic'], {'common':'مشترک','cell-molecular':'سلولی‌مولکولی','microbiology':'میکروبیولوژی'}[q['track']],q['review_in']]
        for d in book for q in d['questions']]))
    h.append(sources_html(all_sources))
    return ''.join(h)


def document_html(title, body, css=STYLE):
    return f'<!doctype html><html lang="fa" dir="rtl"><head><meta charset="utf-8"><title>{html.escape(title)}</title><style>{css}</style></head><body>{body}</body></html>'


def cover_html(volume, back=False):
    if back:
        if volume in (12, 13):
            book = load_book(volume)
            entries = [f'{fa(c["number"])}. {c["title"]}' for c in book['chapters']]
        else:
            entries = ['بیست سؤال منتخب از سال ۱۴۰۳', 'بیست سؤال منتخب از سال ۱۴۰۴', 'بیست سؤال منتخب از سال ۱۴۰۵',
                       'تحلیل مستقل پاسخ و گزینه‌ها', 'منبع و محل مرور هر سؤال', 'نمایه موضوعی برای دو رشته']
        body = '<div class="cover back"><div class="backpanel"><h2>مباحث این جلد</h2>' + ''.join(para(x) for x in entries) + '</div><div class="badge">نسخه آزمایشی غیرقابل انتشار</div></div>'
    else:
        body = (f'<div class="cover"><div class="pill">آمادگی آزمون دکتری ۱۴۰۶</div><div class="pill" style="margin-top:9mm;background:#cf3d52">جلد {fa(volume)}</div>'
                f'<h1>{TITLES[volume]}</h1><div class="subtitle">{("درسنامه گام‌به‌گام و آزمون" if volume != 14 else "سؤال‌های واقعی؛ پاسخ و تحلیل")}</div>'
                '<div class="author">دکتر امیر هوشنگ گرامی</div><div class="badge">نسخه آزمایشی غیرقابل انتشار</div></div>')
    return document_html(TITLES[volume], body, STYLE + '@page{size:A4;margin:0}body{margin:0}html{background:white}')


def rtl_paragraph(p):
    el = OxmlElement('w:bidi'); el.set(qn('w:val'), '1'); p._p.get_or_add_pPr().append(el)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def docx_build(volume, body, cover_png, back_png):
    """Keep prose/questions editable; figures and cover art are embedded images."""
    from lxml import html as lh
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21); section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5); section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(2.5); section.right_margin = Cm(1.5)
    for style in doc.styles:
        if style.type == 1:
            style.font.name = 'B Nazanin'; style.font.size = Pt(14)
            style.paragraph_format.line_spacing = 1.2
            style.element.get_or_add_rPr().rFonts.set(qn('w:cs'), 'B Nazanin')
    settings = doc.settings.element
    settings.append(OxmlElement('w:mirrorMargins'))
    # A source DOCX uses image covers within printable margins; PDF covers are full bleed.
    doc.add_picture(str(cover_png), width=Cm(17))
    doc.add_page_break()
    tree = lh.fromstring('<div>' + body + '</div>')

    def ptext(el, style=None):
        p = doc.add_paragraph(style=style)
        rtl_paragraph(p)
        text = ''.join(el.itertext()).strip()
        for part in re.split(r'([\x20-\x7e]+)', text):
            if not part:
                continue
            r = p.add_run(part)
            r.font.name = 'DejaVu Sans' if re.search('[A-Za-z]', part) else 'B Nazanin'
            rpr = r._r.get_or_add_rPr()
            rtl = OxmlElement('w:rtl'); rtl.set(qn('w:val'), '0' if re.search('[A-Za-z]', part) else '1'); rpr.append(rtl)
        return p

    def walk(el):
        tag = el.tag
        if tag in ('h1', 'h2', 'h3', 'h4'):
            p = ptext(el, 'Heading ' + str(min(int(tag[1]), 3)))
            if 'major' in el.get('class', '') or 'chapter' in el.get('class', ''):
                p.paragraph_format.page_break_before = True
        elif tag == 'p':
            ptext(el)
        elif tag == 'table':
            rows = el.xpath('.//tr')
            t = doc.add_table(rows=0, cols=max(len(row) for row in rows)); t.style = 'Table Grid'
            bidi = OxmlElement('w:bidiVisual'); t._tbl.tblPr.append(bidi)
            for i, row in enumerate(rows):
                cells = t.add_row().cells
                for cell, element in zip(cells, row):
                    cell.text = ''.join(element.itertext())
                    for p in cell.paragraphs: rtl_paragraph(p)
                if i == 0:
                    repeat = OxmlElement('w:tblHeader'); t.rows[-1]._tr.get_or_add_trPr().append(repeat)
        elif tag == 'figure':
            src = el.find('img').get('src')
            path = (OUT / src).resolve().with_suffix('.png')
            doc.add_picture(str(path), width=Cm(16))
            if el.find('figcaption') is not None: ptext(el.find('figcaption'), 'Caption')
        elif tag == 'div' and 'option' in el.get('class', '').split():
            ptext(el)
        elif tag == 'div' and 'citation' in el.get('class', '').split():
            ptext(el)
        elif tag == 'div' and 'source' in el.get('class', '').split():
            ptext(el)
        elif tag == 'div' and 'note' in el.get('class', '').split():
            ptext(el)
        else:
            for child in el: walk(child)
    for child in tree: walk(child)
    doc.add_page_break(); doc.add_picture(str(back_png), width=Cm(17))
    footer = section.footer.paragraphs[0]; footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    field = OxmlElement('w:fldSimple'); field.set(qn('w:instr'), 'PAGE'); footer._p.append(field)
    doc.core_properties.title = f'یاسی — جلد {fa(volume)} — {TITLES[volume]}'
    doc.core_properties.author = 'دکتر امیر هوشنگ گرامی'
    doc.save(OUT / f'Yasi-{volume}.docx')


async def build(volumes):
    OUT.mkdir(exist_ok=True)
    class QuietHandler(SimpleHTTPRequestHandler):
        def log_message(self, *args):
            pass
    server = ThreadingHTTPServer(('127.0.0.1', 0), partial(QuietHandler, directory=str(ROOT)))
    threading.Thread(target=server.serve_forever, daemon=True).start()
    origin = f'http://127.0.0.1:{server.server_port}'
    async with async_playwright() as p:
        browser = await p.chromium.launch(executable_path='/usr/bin/chromium', args=['--no-sandbox'])
        page = await browser.new_page(viewport={'width':794,'height':1123}, device_scale_factor=1.5)
        for volume in volumes:
            book = load_book(volume)
            body = exam_html(book) if volume == 14 else teaching_html(volume, book)
            main = OUT / f'Yasi-{volume}.html'
            main.write_text(document_html(TITLES[volume], body))
            pdf_parts = []
            for side in ('front', 'back'):
                path = OUT / f'Yasi-{volume}-{side}.html'; path.write_text(cover_html(volume, side == 'back'))
                await page.goto(origin + '/output/' + path.name); await page.evaluate('document.fonts.ready')
                await page.screenshot(path=str(OUT / f'Yasi-{volume}-{side}.png'), full_page=True)
                dest = OUT / f'.Yasi-{volume}-{side}.pdf'
                await page.pdf(path=str(dest), prefer_css_page_size=True, print_background=True)
                pdf_parts.append(dest)
            await page.goto(origin + '/output/' + main.name); await page.evaluate('document.fonts.ready')
            await page.emulate_media(media='print')
            main_pdf = OUT / f'.Yasi-{volume}-main.pdf'
            await page.pdf(path=str(main_pdf), prefer_css_page_size=True, print_background=True,
                           outline=True, tagged=True,
                           display_header_footer=True, header_template='<span></span>',
                           footer_template='<div style="width:100%;text-align:center;font-size:10px;color:#4b5c6b"><span class="pageNumber"></span></div>')
            result = fitz.open()
            for path in (pdf_parts[0], main_pdf, pdf_parts[1]):
                with fitz.open(path) as part: result.insert_pdf(part)
            result.set_metadata({'title': f'یاسی — جلد {fa(volume)} — {TITLES[volume]}', 'author':'دکتر امیر هوشنگ گرامی',
                                 'subject':'آمادگی دکتری؛ فقط سلولی‌مولکولی و میکروبیولوژی؛ نسخه آزمایشی'})
            # Preserve useful main-document bookmarks after prepending the front cover.
            with fitz.open(main_pdf) as part:
                toc = part.get_toc()
                if toc: result.set_toc([[level,title,index+1] for level,title,index in toc])
            destination = OUT / f'Yasi-{volume}.pdf'; result.save(destination, garbage=4, deflate=True)
            print(f'Volume {volume}: {len(result)} PDF pages; {destination.stat().st_size:,} bytes', flush=True)
            docx_build(volume, body, OUT / f'Yasi-{volume}-front.png', OUT / f'Yasi-{volume}-back.png')
            for path in (*pdf_parts, main_pdf): path.unlink()
        await browser.close()
    server.shutdown()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('volumes', type=int, nargs='*', default=[12, 13, 14])
    args = parser.parse_args()
    asyncio.run(build(args.volumes))
