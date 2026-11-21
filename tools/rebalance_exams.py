"""Redistribute 100 authored exam questions into four genuinely comprehensive exams."""
import json
import argparse
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def distribute(questions, volume):
    grouped = defaultdict(list)
    for q in questions:
        grouped[q['chapter']].append(q)
    if set(grouped) != set(range(1, 11)) or any(len(v) != 10 for v in grouped.values()):
        raise ValueError('Need ten independent questions from each of ten chapters')
    exams = [dict(number=i, title='جامع همه ده فصل', questions=[]) for i in range(1, 5)]
    for chapter in range(1, 11):
        group = sorted(grouped[chapter], key=lambda q: q.get('origin_id', q['id']))
        sizes = (3, 2, 3, 2) if chapter % 2 else (2, 3, 2, 3)
        offset = 0
        for exam, count in zip(exams, sizes):
            for original in group[offset:offset+count]:
                q = dict(original)
                q['origin_id'] = original.get('origin_id', original['id'])
                q['id'] = f'V{volume}-E{exam["number"]}-Q{len(exam["questions"])+1:02d}'
                exam['questions'].append(q)
            offset += count
    return exams


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('volumes', type=int, nargs='*', default=[12, 13])
    args = parser.parse_args()
    for volume in args.volumes:
        paths = [ROOT / f'volumes/{volume}/part-{p}.json' for p in ('a', 'b')]
        parts = [json.loads(p.read_text()) for p in paths]
        questions = [q for d in parts for e in d['exams'] for q in e['questions']]
        exams = distribute(questions, volume)
        for index, (path, part) in enumerate(zip(paths, parts)):
            part['exams'] = exams[index*2:index*2+2]
            part['integration_note'] = 'بازچینی نهایی آزمون‌ها برای پوشش همه ده فصل؛ شناسه خاستگاه هر سؤال حفظ شد.'
            path.write_text(json.dumps(part, ensure_ascii=False, indent=2) + '\n')
        print(volume, [len(e['questions']) for e in exams])
