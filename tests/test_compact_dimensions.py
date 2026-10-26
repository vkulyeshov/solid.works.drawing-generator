"""Verify the native LeftSide before/after drawing comparison without starting CAD."""
from pathlib import Path
import argparse
import csv
import json
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
import pdfplumber

parser = argparse.ArgumentParser()
parser.add_argument('--before', type=Path, required=True, help='Original external LeftSide output pack')
parser.add_argument('--after', type=Path, required=True, help='Updated external LeftSide output pack')
parser.add_argument('--pattern-pack', type=Path, help='Optional native four-hole plate output pack')
args = parser.parse_args()


def coverage(pack):
    with (pack / 'dimension-coverage.tsv').open(encoding='utf-16', newline='') as stream:
        return list(csv.DictReader(stream, delimiter='\t'))


before = coverage(args.before)
after = coverage(args.after)
folded_sheet = after[0]['Sheet']
before_flat = sum(int(row['Dimensions']) for row in before if row['Sheet'] != before[0]['Sheet'])
after_flat = sum(int(row['Dimensions']) for row in after if row['Sheet'] != folded_sheet)
assert before_flat > 100 and after_flat >= 100, (before_flat, after_flat)
assert sum(int(row['Dimensions']) for row in after if row['Sheet'] == folded_sheet) == 3, after
log = (args.after / 'run.log').read_text('utf-16')
assert 'FAIL ' not in log and 'WARN ' not in log, log[-2000:]
assert 'COORDINATE TABLE' not in log, 'Dimensions were replaced by a text schedule'
routes = re.findall(r'ROUTING .+ axis=(\d+) intervals=(\d+) covered=(\d+)', log)
assert len(routes) == 2 and all(int(expected) == int(covered) for _, expected, covered in routes), routes
assert {int(axis): int(expected) for axis, expected, _ in routes} == {0: 56, 1: 65}, routes
pdf_file = next((args.after / 'PDF').glob('*.pdf'))
with pdfplumber.open(pdf_file) as pdf:
    assert len(pdf.pages) == 2, 'Unexpected coordinate-table sheets'
    removed_marker = '\u0410\u0412\u0422\u041e\u041c\u0410\u0422\u0418\u0427\u041d\u0418\u0419 \u041f\u0420\u041e\u0404\u041a\u0422'
    for page in pdf.pages:
        assert removed_marker not in (page.extract_text() or '')
        assert all(c['x0'] >= -1 and c['x1'] <= page.width + 1 and c['y0'] >= -1 and c['y1'] <= page.height + 1 for c in page.chars), 'Text outside the page'
    assert (pdf.pages[0].extract_text() or '').count('1300,00') == 1, 'Repeated overall length on folded views'
    flat_page = pdf.pages[1]
    dimension_words = [word for word in flat_page.extract_words()
                       if word['top'] * 25.4 / 72 < flat_page.height * 25.4 / 72 - 65
                       and re.fullmatch(r'\d+[,.]\d+', word['text'])]
    assert len(dimension_words) >= 100, 'Graphical dimension labels are missing'
    overlaps = []
    for i, first in enumerate(dimension_words):
        for last in dimension_words[i + 1:]:
            width = min(first['x1'], last['x1']) - max(first['x0'], last['x0'])
            height = min(first['bottom'], last['bottom']) - max(first['top'], last['top'])
            if width > 0 and height > 0 and width * height > 1:
                overlaps.append((first['text'], last['text']))
    assert not overlaps, ('Overlapping dimension text', overlaps)
result = {'passed': True, 'graphic_dimensions_before': before_flat, 'graphic_dimensions_after': after_flat,
          'graphical_intervals_covered': sum(int(covered) for _, _, covered in routes), 'pdf': str(pdf_file.resolve())}
if args.pattern_pack:
    pattern_rows = coverage(args.pattern_pack)
    assert sum(int(row['Dimensions']) for row in pattern_rows) == 4, 'Pitch annotations duplicated overall dimensions'
    pattern_log = (args.pattern_pack / 'run.log').read_text('utf-16')
    assert 'WARN ' not in pattern_log and 'FAIL ' not in pattern_log
    with pdfplumber.open(next((args.pattern_pack / 'PDF').glob('*.pdf'))) as pattern_pdf:
        text = pattern_pdf.pages[0].extract_text() or ''
        assert text.count('4x') == 1, 'Identical diameters were not grouped once'
        assert '5x 30,00 = 150,00' in text, 'Equal pitches were not grouped'
        assert text.count('150,00') == 1, 'Repeated total span'
    result['four_hole_pattern_passed'] = True
report = work_path('compact-dimensions.json')
report.parent.mkdir(parents=True, exist_ok=True)
report.write_text(json.dumps(result, indent=2), encoding='utf8')
print(json.dumps(result, indent=2))
