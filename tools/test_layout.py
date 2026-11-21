"""Regression checks for the editable document's pagination and text fonts."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn
from PIL import Image

import build_books


class WordLayoutTests(unittest.TestCase):
    def test_toc_is_separate_and_short_tables_stay_together(self):
        body = ('<p>مقدمه</p><section class="toc-page"><h2 id="contents">فهرست</h2>'
                '<div class="toc"><p><a href="#chapter-1">فصل ۱</a></p></div></section>'
                '<h2 id="chapter-1">فصل اول</h2>'
                + build_books.table_html(['عنوان', 'عدد'], [['مقدار', '۱۲۳']]))
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cover = root / 'cover.png'
            Image.new('RGB', (20, 28), 'white').save(cover)
            with patch.object(build_books, 'OUT', root):
                build_books.docx_build(12, body, cover, cover)
            doc = Document(root / 'Yasi-12.docx')
            toc = next(p for p in doc.paragraphs if p.text == 'فهرست')
            chapter = next(p for p in doc.paragraphs if p.text == 'فصل اول')
            self.assertTrue(toc.paragraph_format.page_break_before)
            self.assertTrue(chapter.paragraph_format.page_break_before)
            for row in doc.tables[0].rows:
                self.assertIsNotNone(row._tr.find('w:trPr/w:cantSplit', row._tr.nsmap))
            for cell in doc.tables[0].rows[0].cells:
                self.assertTrue(cell.paragraphs[0].paragraph_format.keep_with_next)

    def test_digits_and_latin_in_word_use_vazirmatn_at_body_size(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cover = root / 'cover.png'
            Image.new('RGB', (20, 28), 'white').save(cover)
            with patch.object(build_books, 'OUT', root):
                build_books.docx_build(13, '<p>تیتر ۲۰÷(۰٫۰۵×۱۰⁻³)=۴×۱۰⁵ PFU/mL</p>', cover, cover)
            doc = Document(root / 'Yasi-13.docx')
            for fonts in doc.element.iter(qn('w:rFonts')):
                for key in ('ascii', 'hAnsi', 'cs'):
                    self.assertEqual(fonts.get(qn('w:' + key)), 'Vazirmatn')
            self.assertEqual(doc.styles['Normal'].font.name, 'Vazirmatn')


if __name__ == '__main__':
    unittest.main()
