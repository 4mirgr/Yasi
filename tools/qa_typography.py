"""Check printed table boundaries, contents destinations and typography artifacts."""
import asyncio
import json
import re
import threading
import zipfile
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from xml.etree import ElementTree as ET

import pymupdf as fitz
from lxml import html
from playwright.async_api import async_playwright

from build_books import ROOT, OUT, fa
from qa_outputs import normalized, WORD


def tagged_tables(doc):
    page_numbers = {doc.page_xref(i): i+1 for i in range(len(doc))}

    def pages_below(xref, seen=None):
        seen = set() if seen is None else seen
        if xref in seen: return set()
        seen.add(xref)
        pages = set()
        pg = doc.xref_get_key(xref, 'Pg')[1]
        if pg.endswith(' R'):
            number = page_numbers.get(int(pg.split()[0]))
            if number: pages.add(number)
        children = doc.xref_get_key(xref, 'K')[1]
        for child in re.findall(r'(\d+) 0 R', children):
            ref = int(child)
            if ref in page_numbers: pages.add(page_numbers[ref])
            else: pages.update(pages_below(ref, seen))
        return pages

    tables = []; rows = []; repeated_headers = []
    for xref in range(1, doc.xref_length()):
        role = doc.xref_get_key(xref, 'S')[1]
        if role == '/Table': tables.append(sorted(pages_below(xref)))
        elif role == '/TR':
            children = re.findall(r'(\d+) 0 R', doc.xref_get_key(xref,'K')[1])
            pages = sorted(pages_below(xref))
            is_header = children and all(doc.xref_get_key(int(ref),'S')[1]=='/TH' for ref in children)
            if is_header and len(pages)>1: repeated_headers.append(pages)
            else: rows.append(pages)
    return tables, rows, repeated_headers


async def run():
    class QuietHandler(SimpleHTTPRequestHandler):
        def log_message(self, *args): pass
    server = ThreadingHTTPServer(('127.0.0.1', 0), partial(QuietHandler, directory=str(ROOT)))
    threading.Thread(target=server.serve_forever, daemon=True).start()
    origin = f'http://127.0.0.1:{server.server_port}'
    work = ROOT/'work/qa/typography'; work.mkdir(parents=True, exist_ok=True)
    report = {}
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(executable_path='/usr/bin/chromium', args=['--no-sandbox'])
        page = await browser.new_page()
        await page.emulate_media(media='print')
        for volume in (12,13,14):
            main = OUT/f'Yasi-{volume}.html'
            tree = html.fromstring(main.read_text())
            await page.goto(origin+'/output/'+main.name)
            await page.evaluate('document.fonts.ready')
            body_pdf = work/f'main-{volume}.pdf'
            await page.pdf(path=str(body_pdf), prefer_css_page_size=True, print_background=True, tagged=True)
            body = fitz.open(body_pdf)
            tables, rows, repeated_headers = tagged_tables(body)
            source_tables = tree.xpath('//table')
            assert len(tables) == len(source_tables), (volume, 'tagged table count', len(tables),len(source_tables))
            split_short = [i+1 for i,(el,pages) in enumerate(zip(source_tables,tables))
                           if 'keep-table' in el.get('class','') and len(pages) != 1]
            split_rows = [i+1 for i,pages in enumerate(rows) if len(pages) > 1]
            assert not split_short and not split_rows, (volume, split_short, split_rows)
            doc = fitz.open(OUT/f'Yasi-{volume}.pdf')
            toc_pages = [i for i,p in enumerate(doc) if 'فهرست مطالب' in p.get_text()]
            assert len(toc_pages)==1, (volume,'contents must be one page',toc_pages)
            names = body.resolve_names()
            toc_entries = []
            for link in tree.xpath('//div[@class="toc"]//a'):
                target = link.get('href')[1:]
                label = link.xpath('.//span[@class="toc-number"]')[0].text
                expected = names[target]['page']+2
                assert label == fa(expected), (volume,target,label,expected)
                toc_entries.append({'target':target,'pdf_page':expected})
            assert min(e['pdf_page'] for e in toc_entries)>toc_pages[0]+1
            fonts = {}; bad_digits = []
            for index,p in enumerate(doc):
                for block in p.get_text('dict')['blocks']:
                    for line in block.get('lines',[]):
                        for span in line['spans']:
                            font=span['font'];fonts[font]=fonts.get(font,0)+1
                            if re.search('[0-9۰-۹]',span['text']) and not font.startswith('Vazirmatn'):
                                bad_digits.append({'page':index+1,'text':span['text'],'font':font})
            assert not bad_digits, (volume,'digits in a fallback font',bad_digits)
            with zipfile.ZipFile(OUT/f'Yasi-{volume}.docx') as archive:
                xml = ET.fromstring(archive.read('word/document.xml'))
                word_text = normalized(''.join(t.text or '' for t in xml.iter(WORD+'t')))
                wrong_fonts=[f.attrib for f in xml.iter(WORD+'rFonts')
                             if any(f.get(WORD+k)!='Vazirmatn' for k in ('ascii','hAnsi','cs'))]
                assert not wrong_fonts, (volume,wrong_fonts)
            questions=[]
            for path in sorted((ROOT/f'volumes/{volume}').glob('questions-*.json' if volume==14 else 'part-*.json')):
                data=json.loads(path.read_text())
                if volume==14: questions+=data['questions']
                else: questions += [q for c in data['chapters'] for s in c['steps'] for q in s['questions']]+data['review']+[q for e in data['exams'] for q in e['questions']]
            missing=[]
            for q in questions:
                for name,value in [('stem',q['stem']),('explanation',q['explanation'])]+[(f'option-{i+1}',v) for i,v in enumerate(q['options'])]:
                    if normalized(value) not in word_text: missing.append([q['id'],name])
            assert not missing, (volume,'missing Word fields',missing)
            report[str(volume)]={'pages':len(doc),'contents_page':toc_pages[0]+1,'contents_entries':toc_entries,
                'tables':tables,'table_rows_checked':len(rows),'repeated_table_headers':repeated_headers,
                'split_short_tables':split_short,'split_rows':split_rows,
                'fonts':fonts,'digits_using_fallback':bad_digits,'word_question_fields_checked':len(questions)*6,'word_missing_fields':missing}
            body.close();doc.close()
        await browser.close()
    server.shutdown()
    (ROOT/'docs/qa/typography-checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__=='__main__': asyncio.run(run())
