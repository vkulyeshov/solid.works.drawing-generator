"""Verify role ownership, coordinate completeness and saved native dimensions."""
from pathlib import Path
import argparse
import math
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
import pdfplumber


def number(value):
    return float(value.replace(',', '.'))


def segment_hits_box(line, box):
    lower, upper = 0.0, 1.0
    for axis in (0, 1):
        start, delta = line[axis], line[axis + 2] - line[axis]
        if abs(delta) < 1e-12:
            if not box[axis] <= start <= box[axis + 2]:
                return False
        else:
            a, b = sorted(((box[axis] - start) / delta, (box[axis + 2] - start) / delta))
            lower, upper = max(lower, a), min(upper, b)
            if lower >= upper:
                return False
    return True


def check_pdf_markers(page, saved):
    unit = 25.4 / 72000
    unmatched = [tuple(map(number, row[4:6])) for row in saved if row[0] == 'ARROW']
    markers = 0
    for curve in page.curves:
        if not curve['fill']:
            continue
        points = curve['pts']
        if len(points) == 3:
            tip = points[0][0] * unit, (page.height - points[0][1]) * unit
        elif len(points) > 20 and abs(curve['width'] - curve['height']) < 0.001:
            tip = (curve['x0'] + curve['x1']) * unit / 2, (curve['y0'] + curve['y1']) * unit / 2
        else:
            continue
        assert unmatched, ('Extra PDF arrow or marker', tip)
        index = min(range(len(unmatched)), key=lambda i: math.dist(unmatched[i], tip))
        assert math.dist(unmatched[index], tip) < 0.00002, ('PDF arrow has no native dimension', tip)
        unmatched.pop(index)
        markers += 1
    assert not unmatched, ('Native dimension markers missing from PDF', unmatched)
    assert all(row[8].strip() for row in saved if row[0] == 'DIM'), 'An arrow pair has no dimension text'
    return markers


def check(directory):
    directory = external_path(directory)
    audit = [r.split('\t') for r in (directory / 'dimension-plan.tsv').read_text(encoding='utf-16').splitlines()]
    saved = [r.split('\t') for r in (directory / 'saved-dimensions.tsv').read_text(encoding='utf-16').splitlines()]
    dimensions = {(r[2], r[3]): r for r in saved if r[0] == 'DIM'}
    assert len(audit) == len(dimensions) > 0
    assert len({tuple(r[5:8]) for r in audit}) == len(audit), 'A physical interval was repeated'
    parents = {}

    def root(axis, coordinate):
        key = (axis, coordinate)
        parents.setdefault(key, key)
        while parents[key] != key:
            key = parents[key]
        return key

    views = {}
    for row in audit:
        _, job, view, name, role, axis, low, high, value, side = row
        record = dimensions[(view, name)]
        assert math.isclose(number(record[5]), number(value), abs_tol=0.000005), (row, record)
        assert record[4] == '2' and record[9] == '2', record
        views.setdefault(role, set()).add(view)
        parents[root(axis, low)] = root(axis, high)
    assert all(len(group) == 1 for group in views.values()), views
    outlines = {r[2]: list(map(number, r[3:7])) for r in saved if r[0] == 'VIEW'}
    for axis in ('0', '1'):
        overall = max((row for row in audit if row[4] == 'outline' and row[5] == axis), key=lambda row: number(row[8]))
        arrows = [row for row in saved if row[0] == 'ARROW' and row[2:4] == overall[2:4]]
        assert len(arrows) == 2
        bounds = outlines[overall[2]]
        if axis == '0':
            assert min(number(row[5]) for row in arrows) > bounds[3], 'Overall width must be above the outline view'
        else:
            assert max(number(row[4]) for row in arrows) < bounds[0], 'Overall height must be left of the outline view'
    if 'slot' in views:
        assert views['outline'].isdisjoint(views['slot'])
        assert outlines[next(iter(views['outline']))][1] > outlines[next(iter(views['slot']))][3]

    log = (directory / 'run.log').read_text(encoding='utf-16')
    candidates = re.findall(r'PROFILE PLAN ([01])\|(-?\d+)\|(-?\d+); role=\w+; value=[^;]+; selected=(True|False)', log)
    assert candidates
    for axis, low, high, selected in candidates:
        assert root(axis, low) == root(axis, high), f'Undimensioned coordinate interval: {(axis, low, high)}'
    assert sum(r[3] == 'True' for r in candidates) == len(audit)
    assert len(audit) == len(parents) - len({root(*key) for key in list(parents)})
    assert 'all graphics inside the drawing area' in log
    regions = {}
    for row in (directory / 'native-view-regions.tsv').read_text(encoding='utf-16').splitlines():
        fields = row.split('\t')
        regions[fields[1]] = list(map(number, fields[2:6]))
    for row in saved:
        if row[0] == 'LINE':
            bounds = regions[row[2]]
            points = list(map(number, row[4:8]))
            for x, y in (points[:2], points[2:]):
                assert bounds[0] - 0.0003 <= x <= bounds[2] + 0.0003, row
                assert bounds[1] - 0.0003 <= y <= bounds[3] + 0.0003, row
    pdfs = list(directory.glob('PDF/*.pdf'))
    assert len(pdfs) == 1
    with pdfplumber.open(pdfs[0]) as pdf:
        assert len(pdf.pages) == 1
        lines = [list(map(number, row[4:8])) for row in saved if row[0] == 'LINE']
        for page in pdf.pages:
            markers = check_pdf_markers(page, saved)
            for char in page.chars:
                assert 0 <= char['x0'] <= char['x1'] <= page.width
                assert 0 <= char['top'] <= char['bottom'] <= page.height
                if not char['text'].strip():
                    continue
                dx = (char['x1'] - char['x0']) * 0.22
                dy = (char['bottom'] - char['top']) * 0.25
                unit = 25.4 / 72000
                ink = [(char['x0'] + dx) * unit, (page.height - char['bottom'] + dy) * unit,
                       (char['x1'] - dx) * unit, (page.height - char['top'] - dy) * unit]
                assert not any(segment_hits_box(line, ink) for line in lines), ('Line crosses printed text', char['text'], ink)
    print(f'PASS: {len(audit)} native dimensions; {markers} PDF markers matched to labeled native dimensions; {len(candidates)} source intervals fully constrained; roles={sorted(views)}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', required=True, type=Path)
    check(parser.parse_args().directory)
