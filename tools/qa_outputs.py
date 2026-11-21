"""Check generated artifacts and render every PDF page into numbered contact sheets.

Contact sheets support human review; geometry/text checks do not replace it.
"""
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pymupdf as fitz
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
SVG = '{http://www.w3.org/2000/svg}'
WORD = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def normalized(text):
    from build_books import SCRIPT_CHARS
    text = re.sub(r'(?<=[0-9۰-۹])\^([+−\-]?[0-9۰-۹]+)', r'\1', text)
    text = text.translate(SCRIPT_CHARS)
    return re.sub(r'[\s\u200e\u200f\u2066-\u2069]', '', text)


def run():
    if tuple(int(n) for n in fitz.VersionBind.split('.')[:3]) < (1, 28, 2):
        raise RuntimeError('Install PyMuPDF >= 1.28.2 from tools/requirements.txt; 1.26.6 crashes during extraction of some RTL text.')
    report = {'pdfs': {}, 'docx': {}, 'svg_geometry': {}, 'limits': [
        'Contact sheets require human inspection; machine checks alone do not establish visual or scientific correctness.']}
    report['pymupdf_version'] = fitz.VersionBind
    for volume in (12, 13, 14):
        path = ROOT / f'output/Yasi-{volume}.pdf'
        pdf = fitz.open(path)
        empty = []; outside = []; replacement = []
        dest = ROOT / f'work/qa/{volume}'; dest.mkdir(parents=True, exist_ok=True)
        for pattern in ('sheet-*.png', 'page-*.png'):
            for old in dest.glob(pattern): old.unlink()
        for start in range(0, len(pdf), 16):
            sheet = Image.new('RGB', (1200, 1780), '#e2e8ed'); draw = ImageDraw.Draw(sheet)
            for index in range(start, min(start+16, len(pdf))):
                page = pdf[index]
                text = page.get_text()
                if len(text.strip()) < 15 and index not in (0, len(pdf)-1): empty.append(index+1)
                if '\ufffd' in text: replacement.append(index+1)
                for block in page.get_text('dict')['blocks']:
                    for line in block.get('lines', []):
                        for span in line['spans']:
                            x0,y0,x1,y1 = span['bbox']
                            if x0 < -1 or y0 < -1 or x1 > page.rect.width+1 or y1 > page.rect.height+1:
                                outside.append({'page':index+1,'text':span['text'][:60],'bbox':span['bbox']})
                pix = page.get_pixmap(matrix=fitz.Matrix(.48,.48), alpha=False)
                im = Image.frombytes('RGB', (pix.width,pix.height), pix.samples)
                im.thumbnail((290,411))
                x=(index-start)%4*300; y=(index-start)//4*445
                sheet.paste(im,(x,y+24));draw.text((x+6,y+5),f'PDF page {index+1}',fill='black')
            sheet.save(dest / f'sheet-{start//16+1:02d}.png')
        # Full resolution samples include title, early lesson, middle and ending.
        for index in sorted({0,2,4,6,len(pdf)//2,len(pdf)-3,len(pdf)-1}):
            pdf[index].get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(dest / f'page-{index+1:03d}.png')
        report['pdfs'][str(volume)] = dict(pages=len(pdf), bytes=path.stat().st_size,
            sha256=hashlib.sha256(path.read_bytes()).hexdigest(), empty_body_pages=empty,
            text_outside_page=outside, replacement_character_pages=replacement,
            outline_entries=len(pdf.get_toc()), all_pages_rendered=True,
            contact_sheets=f'work/qa/{volume}/')
        internal = [link for page in pdf for link in page.get_links() if link['kind'] == fitz.LINK_GOTO]
        expected = (ROOT / f'output/Yasi-{volume}.html').read_text().count('href="#')
        report['pdfs'][str(volume)]['internal_links'] = len(internal)
        report['pdfs'][str(volume)]['internal_anchor_elements'] = expected
        report['pdfs'][str(volume)]['bad_link_targets'] = [link['page'] for link in internal if not 0 <= link['page'] < len(pdf)]
        docx = ROOT / f'output/Yasi-{volume}.docx'
        with zipfile.ZipFile(docx) as z:
            xml = ET.fromstring(z.read('word/document.xml'))
            text = normalized(''.join(t.text or '' for t in xml.iter(WORD+'t')))
            links = len(list(xml.iter(WORD+'hyperlink')))
            bookmarks = len(list(xml.iter(WORD+'bookmarkStart')))
            pictures = sum(n.startswith('word/media/') for n in z.namelist())
        paths = sorted((ROOT / f'volumes/{volume}').glob('questions-*.json' if volume == 14 else 'part-*.json'))
        questions=[]
        for p in paths:
            d=json.loads(p.read_text())
            if volume==14:questions+=d['questions']
            else:questions += [q for c in d['chapters']for s in c['steps']for q in s['questions']] + d['review'] + [q for e in d['exams']for q in e['questions']]
        missing = [q['id'] for q in questions if normalized(q['stem']) not in text]
        report['docx'][str(volume)] = dict(bytes=docx.stat().st_size,questions_checked=len(questions),
            missing_question_stems=missing,internal_hyperlinks=links,bookmarks=bookmarks,embedded_images=pictures,
            sha256=hashlib.sha256(docx.read_bytes()).hexdigest())
    failures=[]
    for path in (ROOT/'figures').glob('*/*.svg'):
        root=ET.parse(path).getroot();g=root.find(SVG+'g');polygon=g.find(SVG+'polygon')
        scale=float(re.search(r'scale\(([^ ]+)',g.attrib['transform']).group(1))
        vb=[float(v)for v in root.attrib['viewBox'].split()]
        pts=[tuple(float(v)for v in p.split(','))for p in polygon.attrib['points'].split()]
        width=(max(x for x,y in pts)-min(x for x,y in pts))*scale
        height=(max(y for x,y in pts)-min(y for x,y in pts))*scale
        if width>vb[2]*1.025 or height>vb[3]*1.025:failures.append(str(path.relative_to(ROOT)))
    report['svg_geometry']={'figures_checked':len(list((ROOT/'figures').glob('*/*.svg'))),'clipping_failures':failures}
    (ROOT/'docs/qa/output-checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps(report,ensure_ascii=False,indent=2))
    assert not failures, 'SVG clipping'
    assert all(not v['missing_question_stems'] for v in report['docx'].values()), 'DOCX missing questions'
    assert all(not v['empty_body_pages'] and not v['text_outside_page'] and not v['replacement_character_pages'] for v in report['pdfs'].values()), 'PDF geometry/text failure'
    assert all(v['internal_links'] >= v['internal_anchor_elements'] and not v['bad_link_targets'] for v in report['pdfs'].values()), 'PDF internal anchors lost during merging'


if __name__=='__main__': run()
