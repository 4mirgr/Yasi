"""Structural checks for authored sources; scientific review is separate."""
import argparse
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def validate_question(q, sources):
    label = q.get('id', 'missing ID')
    require(bool(q.get('id')) and bool(q.get('stem')), f'{label}: missing question')
    options = q.get('options', [])
    require(len(options) == 4 and all(isinstance(x, str) and x.strip() for x in options),
            f'{label}: exactly four nonempty options required')
    require(len(set(x.strip() for x in options)) == 4, f'{label}: duplicate options')
    require(type(q.get('answer')) is int and 0 <= q['answer'] < 4, f'{label}: invalid answer index')
    require(bool(q.get('explanation')), f'{label}: missing explanation')
    require(q.get('sources') and set(q['sources']) <= sources, f'{label}: unresolved sources')


def validate_diagram(d):
    nodes = d.get('nodes', [])
    require(d.get('title') and len(nodes) >= 2, 'diagram needs title and nodes')
    require(all(n.get('fa') and n.get('en') for n in nodes), 'diagram labels must be bilingual')
    require(bool(d.get('edges')), 'diagram needs meaningful connections')
    for edge in d['edges']:
        require(len(edge) in (2, 3) and all(type(i) is int and 0 <= i < len(nodes) for i in edge[:2]),
                'diagram edge references absent node')


def validate_all(partial=False):
    report = {'educational': {}, 'exams': {}, 'limitations': [
        'Structural checks do not certify scientific accuracy or official answer keys.']}
    all_ids = set()
    for volume in (12, 13):
        paths = [ROOT / 'volumes' / str(volume) / f'part-{part}.json' for part in ('a', 'b')]
        if partial and not all(p.exists() for p in paths):
            continue
        data = [json.loads(p.read_text()) for p in paths]
        chapters = sum((d['chapters'] for d in data), [])
        require([c['number'] for c in chapters] == list(range(1, 11)), f'V{volume}: chapter coverage')
        steps = sum((c['steps'] for c in chapters), [])
        require(all(len(c['steps']) == 4 for c in chapters), f'V{volume}: steps per chapter')
        require([s['number'] for s in steps] == list(range(1, 41)), f'V{volume}: step coverage')
        source_rows = sum((d['sources'] for d in data), [])
        refs = {s['id'] for s in source_rows}
        require(len(refs) == len(source_rows), f'V{volume}: duplicate reference IDs')
        for s in source_rows:
            require(s.get('url', '').startswith('https://') and s.get('section') and s.get('coverage'),
                    f"{s.get('id')}: incomplete evidence record")
        questions = []
        for c in chapters:
            require(c.get('summary') and c.get('table'), f"V{volume} chapter {c['number']}: summary/table")
            for s in c['steps']:
                require(len(' '.join(s['lesson']).split()) >= 140, f"V{volume} S{s['number']}: outline rather than lesson")
                require(s.get('example', {}).get('problem') and s['example'].get('solution'), 'missing worked example')
                require(s.get('sources') and set(s['sources']) <= refs, 'unresolved lesson sources')
                validate_diagram(s['diagram'])
                require(len(s['questions']) == 3, 'each step needs three questions')
                questions.extend(s['questions'])
        review = sum((d['review'] for d in data), [])
        exams = sum((d['exams'] for d in data), [])
        require(len(review) == 50, f'V{volume}: review must contain 50 independent questions')
        require([e['number'] for e in exams] == [1, 2, 3, 4], f'V{volume}: four exam sequence')
        require(all(len(e['questions']) == 25 for e in exams), f'V{volume}: exam length')
        questions += review + sum((e['questions'] for e in exams), [])
        stems = set()
        answer_counts = [0] * 4
        for q in questions:
            validate_question(q, refs)
            require(q['id'] not in all_ids, f"duplicate ID {q['id']}")
            require(q['stem'].strip() not in stems, f"repeated question stem {q['id']}")
            all_ids.add(q['id']); stems.add(q['stem'].strip())
            answer_counts[q['answer']] += 1
        report['educational'][str(volume)] = dict(chapters=len(chapters), steps=len(steps),
            figures=len(steps), worked_examples=len(steps), step_questions=120,
            review_questions=len(review), exam_questions=100, total_questions=len(questions),
            answer_distribution=answer_counts, sources=len(refs))
    for year in (1403, 1404, 1405):
        path = ROOT / 'volumes' / '14' / f'questions-{year}.json'
        if partial and not path.exists():
            continue
        d = json.loads(path.read_text())
        refs = {s['id'] for s in d['sources']}
        tracks = {}
        for q in d['questions']:
            validate_question(q, refs)
            require(q['id'] not in all_ids, f"duplicate ID {q['id']}")
            require(1 <= q['pdf_page'] <= d['booklet']['pages'], f"{q['id']}: page outside original PDF")
            require(q['track'] in ('common', 'cell-molecular', 'microbiology'), f"{q['id']}: outside requested tracks")
            require((ROOT / q['image']).is_file(), f"{q['id']}: missing image evidence")
            require(q['key']['status'] in ('unknown', 'transcribed-unverified', 'verified'), 'invalid key status')
            require(q.get('review_in'), 'missing educational cross-reference')
            all_ids.add(q['id'])
            tracks[q['track']] = tracks.get(q['track'], 0) + 1
        require(len(d['questions']) >= 20, f'{year}: selected set incomplete')
        require(set(tracks) == {'common', 'cell-molecular', 'microbiology'}, f'{year}: track coverage missing')
        report['exams'][str(year)] = {'questions': len(d['questions']), 'tracks': tracks}
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--partial', action='store_true')
    args = parser.parse_args()
    result = validate_all(args.partial)
    out = ROOT / 'docs' / 'qa' / 'structural-checks.json'
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(json.dumps(result, ensure_ascii=False, indent=2))
