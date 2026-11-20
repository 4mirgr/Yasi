import copy
import unittest

from validate_books import validate_question, validate_diagram
from build_books import esc


class ContentValidationTests(unittest.TestCase):
    def setUp(self):
        self.question = dict(id='V12-S01-Q1', stem='سؤال مستقل',
                             options=['یک', 'دو', 'سه', 'چهار'], answer=2,
                             explanation='توضیح علت پاسخ و رد گزینه‌ها', sources=['R1'])

    def test_answer_must_reference_one_of_four_original_options(self):
        validate_question(self.question, {'R1'})
        for bad in (-1, 4, True, '2'):
            q = copy.deepcopy(self.question)
            q['answer'] = bad
            with self.assertRaises(ValueError):
                validate_question(q, {'R1'})

    def test_missing_or_unresolved_evidence_is_rejected(self):
        for refs in ([], ['NOT-READ']):
            q = copy.deepcopy(self.question)
            q['sources'] = refs
            with self.assertRaises(ValueError):
                validate_question(q, {'R1'})

    def test_duplicate_options_are_rejected(self):
        q = copy.deepcopy(self.question)
        q['options'] = ['یک', 'دو', 'دو', 'چهار']
        with self.assertRaises(ValueError):
            validate_question(q, {'R1'})

    def test_diagram_edge_must_resolve_bilingual_nodes(self):
        d = dict(title='طرح', nodes=[dict(fa='آغاز', en='Start'),
                                    dict(fa='پایان', en='End')], edges=[[0, 1]])
        validate_diagram(d)
        d['edges'] = [[0, 2]]
        with self.assertRaises(ValueError):
            validate_diagram(d)

    def test_formula_is_isolated_without_corrupting_html_entities(self):
        self.assertIn('<bdi>5′→3′</bdi>', esc('جهت 5′→3′ است'))
        self.assertIn('<bdi>ΔG &lt; 0</bdi>', esc('شرط ΔG < 0'))
        self.assertNotIn('&<bdi>', esc('شرط ΔG < 0'))


if __name__ == '__main__':
    unittest.main()
