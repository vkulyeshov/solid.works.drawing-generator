"""Verify the external Matochina4x100 native postprocessing result."""
import json
from pathlib import Path
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
import pdfplumber

directory = work_path('hub_leaders/comparison')
report = json.loads((directory / 'postprocess-comparison.json').read_text('utf-8-sig'))
assert len(report) == 1 and report[0]['mode'] == 'postprocessed_chain'
result = report[0]
assert result['valuesUnchanged']
assert 60 <= result['originalDimensions'] <= 66, result['originalDimensions']
assert 20 <= result['dimensions'] <= 50, result['dimensions']
assert '_failed' not in Path(result['pdf']).stem
regions = {}
for line in work_path('hub_leaders/native-view-regions.tsv').read_text('utf-16').splitlines():
    fields = line.split('\t')
    regions[fields[1]] = [float(value.replace(',', '.')) for value in fields[2:6]]
assert len(regions) == 4
owners = {dimension['view'] for dimension in result['dimensionsData']}
for name, first in regions.items():
    if name not in owners:
        continue
    for other, last in regions.items():
        if name == other or other not in owners:
            continue
        assert first[2] <= last[0] or last[2] <= first[0] or first[3] <= last[1] or last[3] <= first[1], (name, other)
audit = [line.split('\t') for line in work_path('hub_leaders/native-dimension-audit.tsv').read_text('utf-16').splitlines()]
assert len(audit) == result['originalDimensions'] - result['dimensions']
assert any('axis-interval|2|' in row[4] for row in audit), 'No mirrored model-Z interval was removed'
names = {d['name'] for d in result['dimensionsData']}
assert all(row[2] not in names and row[3] in names for row in audit)
original = json.loads((directory / 'input-geometry.json').read_text('utf-8-sig'))
source_dimensions = {dimension['name']: dimension for dimension in original['dimensions']}
assert len(source_dimensions) == result['originalDimensions']
for row in audit:
    removed = source_dimensions[row[2]]
    retained = source_dimensions[row[3]]
    assert abs(removed['value'] - retained['value']) < 1e-9
    if row[4].startswith('axis-interval|'):
        fields = row[4].split('|')
        axis = int(fields[1])
        low, high = fields[2].split(':')
        assert float(low) != float(high), 'A zero projection was used as a measurement key'
        def interval(dimension):
            first, last = sorted(round(point[axis] * 1e6) for point in dimension['references'][:2])
            if axis == 2:
                return min((first, last), (-last, -first))
            return first, last
        assert interval(removed) == interval(retained), (row[2], row[3])
linear = [d for d in result['dimensionsData'] if d['type'] == 2]
assert any(d['offsetText'] for d in linear)
centerlines = [row.split('\t') for row in work_path('hub_leaders/native-symmetry-lines.tsv').read_text('utf-16').splitlines()]
assert len(centerlines) == 2
for row in centerlines:
    region = regions[row[1]]
    points = [float(value.replace(',', '.')) for value in row[2:6]]
    for x, y in [(points[0], points[1]), (points[2], points[3])]:
        assert region[0] <= x <= region[2] and region[1] <= y <= region[3]
diagonal_dimensions = []
for dimension in result['dimensionsData']:
    region = regions[dimension['view']]
    for line in dimension['lines']:
        for x, y in [(line[4], line[5]), (line[7], line[8])]:
            assert region[0] - 0.0003 <= x <= region[2] + 0.0003 and region[1] - 0.0003 <= y <= region[3] + 0.0003, (dimension['name'], x, y, region)
    if dimension['type'] == 2 and any(
        abs(line[7] - line[4]) > 0.002 and abs(line[8] - line[5]) > 0.002
        for line in dimension['lines']
    ):
        diagonal_dimensions.append(dimension['name'])
assert diagonal_dimensions, 'No inclined leader remains for the conflicting native labels'
with pdfplumber.open(result['pdf']) as pdf:
    assert len(pdf.pages) == 1
    page = pdf.pages[0]
    assert abs(page.width * 25.4 / 72 - 210) < 0.1
    assert abs(page.height * 25.4 / 72 - 297) < 0.1
    for char in page.chars:
        assert char['x0'] >= 0 and char['x1'] <= page.width
        assert char['top'] >= 0 and char['bottom'] <= page.height
print(f'PASS: {result["dimensions"]} retained dimensions with unchanged values; '
      f'{len(audit)} audited duplicates removed; {len(diagonal_dimensions)} inclined linear leaders; '
      'one A4 sheet; disjoint per-view dimension regions')
