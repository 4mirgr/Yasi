"""Build editable DOCX, linked HTML and A4 PDFs from the reviewed JSON sources.

Requires Python python-docx, playwright, PyMuPDF, system Graphviz and Chromium.
Run from any directory: python tools/build_books.py [12 13 14]
"""
import argparse
import asyncio
import base64
import html
import json
import re
import subprocess
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from cover_art import science_art

import pymupdf as fitz
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
SCRIPT_CHARS = str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ᶜˢᵐ₀₁₂₃₄₅₆₇₈₉₊₋', '0123456789+−=()csm0123456789+−')
LATIN_RE = re.compile(r"[+−\-]?[A-Za-z0-9\u00b0-\u024f\u02e2\u0370-\u03ff\u06f0-\u06f9\u1d00-\u1dbf][A-Za-z0-9\u00b0-\u024f\u02e2\u0370-\u03ff\u06f0-\u06f9\u1d00-\u1dbf\u2070-\u209f\u2200-\u22ff\s._,:;/\\+=<>&!′’'−–\-→←±×()%\[\]{}^٫٬]*")
STYLE = (ROOT / 'assets/styles/book.css').read_text()


def fa(n):
    return str(n).translate(NUMBERS)


def script_markup(escaped):
    escaped = re.sub(r'(?<=[0-9۰-۹])\^([+−\-]?[0-9۰-۹]+)', r'<sup>\1</sup>', escaped)
    escaped = re.sub(r'[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ᶜˢᵐ]+', lambda m:'<sup>'+m.group().translate(SCRIPT_CHARS)+'</sup>', escaped)
    return re.sub(r'[₀₁₂₃₄₅₆₇₈₉₊₋]+', lambda m:'<sub>'+m.group().translate(SCRIPT_CHARS)+'</sub>', escaped)


def esc(value):
    # Isolate Latin runs so formula/identifier order is stable within Persian text.
    value = str(value)
    output = []; end = 0
    for match in LATIN_RE.finditer(value):
        token = match.group().rstrip()
        start = match.start()
        if start > end and value[start-1] in '([' and token.count({'(':')','[':']'}[value[start-1]]) > token.count(value[start-1]):
            start -= 1
            token = value[start] + token
        output.append(html.escape(value[end:start]))
        end = start + len(token)
        token = script_markup(html.escape(token))
        output.append('<bdi>' + token + '</bdi>')
    output.append(html.escape(value[end:]))
    return ''.join(output)


def para(text, cls=''):
    return f'<p class="{cls}">{esc(text)}</p>'


def refs_html(ids):
    return '؛ '.join(f'<a href="#ref-{html.escape(i)}"><bdi>[{html.escape(i)}]</bdi></a>' for i in ids)


def question_html(q, number=None, compact=False):
    label = fa(number) if number is not None else q['id']
    long = max(map(len, q['options'])) > 65 or sum(map(len, q['options'])) > 170
    options = ''.join(f'<div class="option">{fa(i+1)}) {esc(t)}</div>' for i, t in enumerate(q['options']))
    return (f'<div class="question {"exam-question" if compact else ""}" id="{html.escape(q["id"])}"><p><strong>{esc(label)} — {esc(q["stem"])}</strong></p>'
            f'<div class="options {"long" if long else ""}">{options}</div></div>')


def answer_html(q, number=None):
    label = fa(number) if number is not None else q['id']
    qualification = {'ambiguous':' (پاسخ انتظاری؛ صورت دارای ابهام)',
                     'conditional':' (پاسخ مشروط به فرض‌های شرح)'} .get(q.get('answer_status'), '')
    return (f'<div class="answer"><p><strong>{esc(label)} — گزینه {fa(q["answer"]+1)}{qualification}.</strong> '
            f'{esc(q["explanation"])}</p><div class="citation">{refs_html(q["sources"])}</div></div>')


def table_html(headers, rows):
    # Short summaries/answer keys stay whole; multi-page indexes repeat headers
    # and break only between complete rows instead of forcing a giant table.
    cls = 'keep-table' if len(rows) <= 8 else 'flow-table'
    return f'<table class="{cls}"><thead><tr>' + ''.join(f'<th>{esc(t)}</th>' for t in headers) + '</tr></thead><tbody>' + ''.join(
        '<tr>' + ''.join(f'<td>{esc(t)}</td>' for t in row) + '</tr>' for row in rows) + '</tbody></table>'


def contents_html(entries):
    short = ' short' if len(entries) < 8 else ''
    rows = ''.join(f'<p class="toc-entry"><a href="#{html.escape(target)}">'
                   f'<span class="toc-label">{esc(title)}</span>'
                   '<span class="toc-number"></span></a></p>' for target, title in entries)
    return (f'<section class="toc-page{short}"><h2 id="contents">فهرست مطالب</h2>'
            '<p class="toc-caption">برای رفتن به هر بخش، عنوان آن را انتخاب کنید.</p>'
            f'<div class="toc">{rows}</div></section>')


def make_diagram(volume, step):
    """Render authored semantic diagrams with Graphviz/Pango Arabic shaping."""
    directory = ROOT / 'figures' / str(volume)
    directory.mkdir(parents=True, exist_ok=True)
    base = directory / f'step-{step["number"]:02d}'
    d = step['diagram']
    orientation = 'TB' if len(d['nodes']) >= 4 else 'LR'
    lines = ['digraph G {', f'graph [rankdir={orientation}, bgcolor="white", pad="0.2", nodesep="0.3", ranksep="0.2", size="7,3.5", ratio="compress"];',
             'node [shape=box, style="rounded,filled", fillcolor="#eaf5fa", color="#247c91", fontname="Vazirmatn", fontsize=16, margin="0.16,0.12"];',
             'edge [color="#7454a5", arrowsize=0.8, fontname="Vazirmatn", fontsize=12];']
    for i, n in enumerate(d['nodes']):
        label = n['fa'] + '\n' + n['en']
        marked = script_markup(html.escape(label))
        if marked != html.escape(label):
            # Graphviz does not keep an HTML SUP attached to an LTR identifier
            # inside a mixed RTL text node. Give each directional run its own
            # cell; the Latin base and its superscript remain a single unit.
            runs = [s for s in re.split(r'(<bdi>.*?</bdi>)', esc(n['fa'])) if s.strip()]
            cells = [s.replace('<bdi>','').replace('</bdi>','') for s in reversed(runs)]
            marked = ('<TABLE BORDER="0" CELLBORDER="0" CELLPADDING="1"><TR>'
                      + ''.join('<TD>'+s+'</TD>' for s in cells) + '</TR>'
                      + f'<TR><TD COLSPAN="{len(cells)}">{script_markup(html.escape(n["en"]))}</TD></TR></TABLE>')
            marked = marked.replace('<sup>','<SUP><FONT POINT-SIZE="12">').replace('</sup>','</FONT></SUP>').replace('<sub>','<SUB><FONT POINT-SIZE="12">').replace('</sub>','</FONT></SUB>')
            lines.append(f'n{i} [label=<{marked}>];')
        else:
            lines.append(f'n{i} [label={json.dumps(label, ensure_ascii=False)}];')
    for e in d['edges']:
        label = f' [label={json.dumps(e[2], ensure_ascii=False)}]' if len(e) == 3 and e[2] else ''
        lines.append(f'n{e[0]} -> n{e[1]}{label};')
    lines.append('}')
    base.with_suffix('.dot').write_text('\n'.join(lines))
    for fmt in ('svg', 'png'):
        # Graphviz 2.42 scales the SVG graph transform incorrectly when dpi is
        # overridden together with size/ratio. SVG is vector; only PNG needs dpi.
        dpi = ['-Gdpi=180'] if fmt == 'png' else []
        subprocess.run(['dot', '-T' + fmt, *dpi, str(base.with_suffix('.dot')),
                        '-o', str(base.with_suffix('.' + fmt))], check=True, capture_output=True)
        if fmt == 'svg':
            svg_path = base.with_suffix('.svg')
            svg = svg_path.read_text()
            data = base64.b64encode((ROOT/'assets/fonts/Vazirmatn-Regular.ttf').read_bytes()).decode('ascii')
            style = '<style>@font-face{font-family:Vazirmatn;src:url(data:font/ttf;base64,'+data+')}</style>'
            start = svg.index('>',svg.index('<svg'))+1
            svg_path.write_text(svg[:start]+'\n'+style+svg[start:])
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
         '<div class="note">نسخه آزمایشی غیرقابل انتشار. طبق صفحه چاپی ۳۳ دو نسخه بازنشرشده دفترچه ۱۴۰۶ و تصویر ارسالی کاربر، در مجموعه ۲۲۲۶ مسیر سلولی‌مولکولی ترکیب ۱ (دروس ۱،۲،۳،۴) و میکروبیولوژی ترکیب ۴ (دروس ۱،۲،۷) است. جلدهای ۱۲ و ۱۳ بخشی از ماده ۷ مسیر میکروبیولوژی‌اند؛ برای مسیر سلولی‌مولکولی، مبانی مشترک را تقویت می‌کنند و جایگزین مواد اختصاصی ۳ و ۴ نیستند. ماده ۷ بوم‌شناسی میکروارگانیسم‌ها را نیز دارد؛ مطالعه آن را جدا کامل کنید. دریافت مستقیم مبدأ سنجش ممکن نشد؛ علامت انتخاب درس، ضریب عددی زیرمبحث نیست. این جزوه پیش‌بینی سهم سؤال‌های آینده نیست.</div>',
         ]
    entries = [(f'chapter-{c["number"]}', f'فصل {fa(c["number"])} — {c["title"]}') for c in book['chapters']]
    entries += [('review', 'بانک پنجاه سؤال مرور'), ('exams', 'چهار آزمون جامع'),
                ('exam-answers', 'کلید و پاسخ تشریحی چهار آزمون'), ('glossary', 'واژه‌نامه'), ('sources', 'منابع')]
    h.append(contents_html(entries))
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
        cls = 'exam-start' if e['number'] > 1 else ''
        h.append(f'<h2 class="{cls}" id="exam-{e["number"]}">آزمون {fa(e["number"])} — {esc(e["title"])}</h2>')
        h.extend(question_html(q, i+1, compact=True) for i, q in enumerate(e['questions']))
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
         ]
    entries = [(f'year-{d["year"]}', f'سال {fa(d["year"])} — {fa(len(d["questions"]))} سؤال') for d in book]
    entries += [('topic-index', 'نمایه موضوعی و محل مرور'), ('sources', 'منابع علمی')]
    h.append(contents_html(entries))
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
            entries = [f'<p><span>{fa(c["number"])}.</span> {esc(c["title"])}</p>' for c in book['chapters']]
        else:
            entries = ['بیست سؤال منتخب از سال ۱۴۰۳', 'بیست سؤال منتخب از سال ۱۴۰۴', 'بیست سؤال منتخب از سال ۱۴۰۵',
                       'تحلیل مستقل پاسخ و گزینه‌ها', 'منبع و محل مرور هر سؤال', 'نمایه موضوعی برای دو رشته']
            entries = [para(x) for x in entries]
        body = '<div class="cover back"><div class="backpanel"><h2>مباحث این جلد</h2>' + ''.join(entries) + '</div><div class="badge">نسخه آزمایشی غیرقابل انتشار</div></div>'
    else:
        body = (f'<div class="cover v{volume}"><div class="pill">آمادگی آزمون دکتری ۱۴۰۶</div><div class="pill" style="margin-top:9mm;background:#cf3d52">جلد {fa(volume)}</div>'
                f'<h1>{TITLES[volume]}</h1>{science_art(volume)}<div class="subtitle">{("درسنامه گام‌به‌گام و آزمون" if volume != 14 else "سؤال‌های واقعی؛ پاسخ و تحلیل")}</div>'
                '<div class="author">دکتر امیر هوشنگ گرامی</div><div class="badge">نسخه آزمایشی غیرقابل انتشار</div></div>')
    return document_html(TITLES[volume], body, STYLE + '@page{size:A4;margin:0}@page:left{margin:0}@page:right{margin:0}@page:first{margin:0}html,body{margin:0;padding:0;width:210mm;max-width:none;background:white}')


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
            size = {'Heading 1':24, 'Heading 2':19, 'Heading 3':16, 'Caption':10.5}.get(style.name, 13)
            style.font.name = 'Vazirmatn'; style.font.size = Pt(size)
            style.paragraph_format.line_spacing = 1.65
            props = style.element.get_or_add_rPr()
            for key in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
                props.rFonts.set(qn('w:' + key), 'Vazirmatn')
            cs_size = OxmlElement('w:szCs'); cs_size.set(qn('w:val'), str(int(size*2))); props.append(cs_size)
    settings = doc.settings.element
    settings.append(OxmlElement('w:mirrorMargins'))
    # A source DOCX uses image covers within printable margins; PDF covers are full bleed.
    doc.add_picture(str(cover_png), width=Cm(17))
    doc.add_page_break()
    tree = lh.fromstring('<div>' + body + '</div>')
    bookmark_counter = 0
    page_before_next = False

    def bookmark_name(identifier):
        return ('b_' + re.sub('[^A-Za-z0-9_]', '_', identifier))[:40]

    def rich_runs(container, text, bold=False, vertical=None):
        end = 0
        spans = []
        for match in LATIN_RE.finditer(text):
            token = match.group().rstrip()
            start = match.start()
            if start > end and text[start-1] in '([' and token.count({'(':')','[':']'}[text[start-1]]) > token.count(text[start-1]):
                start -= 1
                token = text[start] + token
            spans.append((text[end:start], False))
            spans.append((token, True))
            end = start + len(token)
        spans.append((text[end:], False))
        for part, latin in spans:
            if not part: continue
            run = OxmlElement('w:r'); props = OxmlElement('w:rPr')
            fonts = OxmlElement('w:rFonts')
            fonts.set(qn('w:ascii'), 'Vazirmatn')
            fonts.set(qn('w:hAnsi'), 'Vazirmatn')
            fonts.set(qn('w:cs'), 'Vazirmatn'); fonts.set(qn('w:eastAsia'), 'Vazirmatn'); props.append(fonts)
            rtl = OxmlElement('w:rtl'); rtl.set(qn('w:val'), '0' if latin else '1'); props.append(rtl)
            if bold:
                props.append(OxmlElement('w:b')); props.append(OxmlElement('w:bCs'))
            if vertical:
                align = OxmlElement('w:vertAlign'); align.set(qn('w:val'), vertical); props.append(align)
            run.append(props); content = OxmlElement('w:t'); content.set(qn('xml:space'), 'preserve')
            content.text = ('\u200e' + part + '\u200e') if latin else part
            run.append(content); container.append(run)

    def inline(container, element, bold=False, vertical=None):
        vertical = {'sup':'superscript', 'sub':'subscript'}.get(element.tag, vertical)
        if element.text: rich_runs(container, element.text, bold, vertical)
        for child in element:
            href = child.get('href', '')
            if child.tag == 'a' and href.startswith('#'):
                link = OxmlElement('w:hyperlink'); link.set(qn('w:anchor'), bookmark_name(href[1:]))
                inline(link, child, bold, vertical); container.append(link)
            else:
                inline(container, child, bold or child.tag == 'strong', vertical)
            if child.tail: rich_runs(container, child.tail, bold, vertical)

    def ptext(el, style=None):
        nonlocal bookmark_counter, page_before_next
        p = doc.add_paragraph(style=style)
        rtl_paragraph(p)
        if page_before_next:
            p.paragraph_format.page_break_before = True
            page_before_next = False
        if 'toc-entry' in el.get('class', ''):
            # Word multiplies the font's natural line box, not the nominal point
            # size. Vazirmatn's tall metrics would push 15 entries past one page.
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.line_spacing = Pt(24)
        if el.get('id'):
            bookmark_counter += 1
            start = OxmlElement('w:bookmarkStart'); start.set(qn('w:id'), str(bookmark_counter))
            start.set(qn('w:name'), bookmark_name(el.get('id'))); p._p.append(start)
        inline(p._p, el)
        if el.get('id'):
            end = OxmlElement('w:bookmarkEnd'); end.set(qn('w:id'), str(bookmark_counter)); p._p.append(end)
        return p

    def walk(el):
        nonlocal page_before_next
        tag = el.tag
        if tag == 'section' and 'toc-page' in el.get('class', ''):
            first_paragraph = len(doc.paragraphs)
            page_before_next = True
            for child in el: walk(child)
            title = doc.paragraphs[first_paragraph]
            title.paragraph_format.line_spacing = Pt(38)
            title.paragraph_format.space_before = Pt(0)
            title.paragraph_format.space_after = Pt(12)
            for paragraph in doc.paragraphs[first_paragraph+1:]:
                paragraph.paragraph_format.space_before = Pt(0)
                if not paragraph._p.xpath('.//w:hyperlink'):
                    paragraph.paragraph_format.line_spacing = Pt(24)
                    paragraph.paragraph_format.space_after = Pt(12)
            page_before_next = True
        elif tag in ('h1', 'h2', 'h3', 'h4'):
            p = ptext(el, 'Heading ' + str(min(int(tag[1]), 3)))
            if 'exam-start' in el.get('class', ''):
                p.paragraph_format.page_break_before = True
        elif tag == 'p':
            ptext(el)
        elif tag == 'strong':
            ptext(el, 'Heading 3')
        elif tag == 'table':
            rows = el.xpath('.//tr')
            t = doc.add_table(rows=0, cols=max(len(row) for row in rows)); t.style = 'Table Grid'
            bidi = OxmlElement('w:bidiVisual'); t._tbl.tblPr.append(bidi)
            for i, row in enumerate(rows):
                cells = t.add_row().cells
                row_props = t.rows[-1]._tr.get_or_add_trPr()
                row_props.append(OxmlElement('w:cantSplit'))
                for cell, element in zip(cells, row):
                    cell.text = ''
                    for p in cell.paragraphs:
                        rtl_paragraph(p)
                        inline(p._p, element, i == 0)
                        p.paragraph_format.space_after = Pt(5)
                        p.paragraph_format.keep_together = True
                        if 'keep-table' in el.get('class', '') and i < len(rows)-1:
                            p.paragraph_format.keep_with_next = True
                if i == 0:
                    repeat = OxmlElement('w:tblHeader'); t.rows[-1]._tr.get_or_add_trPr().append(repeat)
        elif tag == 'figure':
            from PIL import Image
            src = el.find('img').get('src')
            path = (OUT / src).resolve().with_suffix('.png')
            with Image.open(path) as im: width, height = im.size
            scale = min(16/width, 9/height)
            doc.add_picture(str(path), width=Cm(width*scale), height=Cm(height*scale))
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
        await page.emulate_media(media='print')
        for volume in volumes:
            book = load_book(volume)
            body = exam_html(book) if volume == 14 else teaching_html(volume, book)
            main = OUT / f'Yasi-{volume}.html'
            main.write_text(document_html(TITLES[volume], body))
            pdf_parts = []
            for side in ('front', 'back'):
                path = OUT / f'Yasi-{volume}-{side}.html'; path.write_text(cover_html(volume, side == 'back'))
                await page.goto(origin + '/output/' + path.name); await page.evaluate('document.fonts.ready')
                # Never silently deliver a clipped one-line cover title.
                if side == 'front':
                    title_fits = await page.locator('.cover h1').evaluate('''el => {
                        const range = document.createRange(); range.selectNodeContents(el);
                        return range.getBoundingClientRect().width <= el.clientWidth + 1;
                    }''')
                    if not title_fits:
                        raise ValueError(f'Volume {volume}: cover title exceeds its line')
                clipped = await page.locator('.cover').evaluate('''cover => {
                    const bounds = cover.getBoundingClientRect();
                    return [...cover.querySelectorAll('.badge,.backpanel,.backpanel p,.author')]
                        .filter(el => {const r=el.getBoundingClientRect();
                            return r.bottom>bounds.bottom-12 || r.left<bounds.left || r.right>bounds.right;})
                        .map(el=>el.textContent);
                }''')
                if clipped:
                    raise ValueError(f'Volume {volume}: clipped {side} cover text: {clipped}')
                await page.screenshot(path=str(OUT / f'Yasi-{volume}-{side}.png'), full_page=True)
                dest = OUT / f'.Yasi-{volume}-{side}.pdf'
                await page.pdf(path=str(dest), prefer_css_page_size=True, print_background=True)
                with fitz.open(dest) as cover:
                    if len(cover) != 1:
                        raise ValueError(f'Volume {volume}: {side} cover is not one page')
                pdf_parts.append(dest)
            await page.goto(origin + '/output/' + main.name); await page.evaluate('document.fonts.ready')
            await page.emulate_media(media='print')
            main_pdf = OUT / f'.Yasi-{volume}-main.pdf'
            pdf_settings = dict(prefer_css_page_size=True, print_background=True, outline=True, tagged=True)
            await page.pdf(path=str(main_pdf), **pdf_settings)
            # Resolve the actual printed destinations, including the front-cover
            # offset. Reserved number columns keep the second print's layout fixed.
            with fitz.open(main_pdf) as first_pass:
                names = first_pass.resolve_names()
                entries = await page.locator('.toc a').evaluate_all('(els) => els.map(el => el.hash.slice(1))')
                numbers = {key:fa(names[key]['page']+2) for key in entries}
                first_pages = len(first_pass)
            await page.evaluate('''numbers => {
                for (const link of document.querySelectorAll('.toc a'))
                    link.querySelector('.toc-number').textContent = numbers[link.hash.slice(1)];
            }''', numbers)
            main.write_text(await page.content())
            await page.pdf(path=str(main_pdf), **pdf_settings)
            with fitz.open(main_pdf) as second_pass:
                if len(second_pass) != first_pages or any(
                        fa(second_pass.resolve_names()[key]['page']+2) != number for key, number in numbers.items()):
                    raise ValueError('Contents numbers changed pagination')
            result = fitz.open()
            for path in (pdf_parts[0], main_pdf, pdf_parts[1]):
                with fitz.open(path) as part: result.insert_pdf(part)
            result.set_metadata({'title': f'یاسی — جلد {fa(volume)} — {TITLES[volume]}', 'author':'دکتر امیر هوشنگ گرامی',
                                 'subject':'آمادگی دکتری؛ فقط سلولی‌مولکولی و میکروبیولوژی؛ نسخه آزمایشی'})
            # Preserve useful main-document bookmarks after prepending the front cover.
            with fitz.open(main_pdf) as part:
                toc = part.get_toc()
                if toc: result.set_toc([[level,title,index+1] for level,title,index in toc])
                # insert_pdf does not copy Chromium's /Dests name tree. Resolve
                # named anchors into explicit page destinations before saving.
                for page_index, source_page in enumerate(part):
                    for link in source_page.get_links():
                        if link['kind'] != fitz.LINK_NAMED or link.get('page', -1) < 0:
                            continue
                        target = link['page']
                        point = link.get('to', fitz.Point(0, part[target].rect.height))
                        point = point * part[target].transformation_matrix
                        result[page_index+1].insert_link({'kind':fitz.LINK_GOTO, 'from':link['from'],
                            'page':target+1, 'to':point, 'zoom':link.get('zoom', 0)})
            destination = OUT / f'Yasi-{volume}.pdf'
            temporary = OUT / f'.Yasi-{volume}-assembled.pdf'
            # Chromium's header/footer template does not load document webfonts.
            # Insert readable Persian page numbers using the bundled font itself.
            font_path = ROOT / 'assets/fonts/Vazirmatn-Regular.ttf'
            footer_font = fitz.Font(fontfile=str(font_path))
            for index in range(1, len(result)-1):
                target_page = result[index]
                label = fa(index+1)
                width = footer_font.text_length(label, fontsize=10)
                target_page.insert_text(((target_page.rect.width-width)/2, target_page.rect.height-22),
                    label, fontname='VazirmatnFooter', fontfile=str(font_path), fontsize=10, color=(.26,.36,.42))
            result.save(temporary, garbage=4, deflate=True)
            temporary.replace(destination)
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
