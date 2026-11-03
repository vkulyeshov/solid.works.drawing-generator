"""Audit reopened dimension graphics against the actual projected profile material."""
from pathlib import Path
from collections import defaultdict
import argparse
import math
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
import pdfplumber

TOLERANCE = 0.000002
PDF_UNIT = 25.4 / 72000


def number(value):
    return float(value.replace(',', '.'))


def edges(polygon):
    return zip(polygon, polygon[1:] + polygon[:1])


def location(point, polygon):
    inside = False
    x, y = point
    for (ax, ay), (bx, by) in edges(polygon):
        cross = (x - ax) * (by - ay) - (y - ay) * (bx - ax)
        if abs(cross) <= TOLERANCE * math.hypot(bx - ax, by - ay):
            if min(ax, bx) - TOLERANCE <= x <= max(ax, bx) + TOLERANCE and min(ay, by) - TOLERANCE <= y <= max(ay, by) + TOLERANCE:
                return -1
        if (ay > y) != (by > y) and x < ax + (y - ay) * (bx - ax) / (by - ay):
            inside = not inside
    return int(inside)


def cuts_material(first, last, loops):
    dx, dy = last[0] - first[0], last[1] - first[1]
    if math.hypot(dx, dy) <= TOLERANCE:
        return False
    parameters = {0.0, 1.0}
    for polygon in loops:
        for (ax, ay), (bx, by) in edges(polygon):
            ex, ey = bx - ax, by - ay
            denominator = dx * ey - dy * ex
            if abs(denominator) <= TOLERANCE ** 2:
                continue
            t = ((ax - first[0]) * ey - (ay - first[1]) * ex) / denominator
            u = ((ax - first[0]) * dy - (ay - first[1]) * dx) / denominator
            if 0 < t < 1 and -1e-8 <= u <= 1 + 1e-8:
                parameters.add(t)
    parameters = sorted(parameters)
    for low, high in zip(parameters, parameters[1:]):
        if high - low > 1e-10:
            mid = (low + high) / 2
            point = first[0] + mid * dx, first[1] + mid * dy
            if location(point, loops[0]) == 1 and all(location(point, hole) == 0 for hole in loops[1:]):
                return True
    return False


def records(directory):
    return [row.split('\t') for row in (directory / 'saved-dimensions.tsv').read_text(encoding='utf-16').splitlines()]


def crossings(rows, contours, witnesses=None):
    failures = []
    witnesses = witnesses or {}
    for row in rows:
        if row[0] == 'LINE' and row[2] in contours:
            values = list(map(number, row[4:8]))
            if cuts_material(values[:2], values[2:], contours[row[2]]):
                witness = witnesses.get((row[2], row[3]))
                if witness:
                    axis, a, b = witness
                    if abs(values[axis] - values[axis + 2]) <= TOLERANCE and min(abs(values[axis] - a[axis]), abs(values[axis] - b[axis])) <= TOLERANCE:
                        continue
                failures.append(row)
    return failures


def point_distance(point, first, last):
    dx, dy = last[0] - first[0], last[1] - first[1]
    denominator = dx * dx + dy * dy
    t = max(0, min(1, ((point[0] - first[0]) * dx + (point[1] - first[1]) * dy) / denominator)) if denominator else 0
    return math.hypot(point[0] - first[0] - t * dx, point[1] - first[1] - t * dy)


def check(directory, before=None):
    directory = external_path(directory)
    polygons = defaultdict(lambda: defaultdict(list))
    for line in (directory / 'profile-contours.tsv').read_text(encoding='utf-16').splitlines():
        job, view, loop, index, x, y = line.split('\t')
        polygons[view][int(loop)].append((number(x), number(y)))
    contours = {view: [loops[index] for index in sorted(loops)] for view, loops in polygons.items()}
    assert contours
    saved = records(directory)
    witnesses = {}
    witness_file = directory / 'profile-witnesses.tsv'
    if witness_file.exists():
        for row in witness_file.read_text(encoding='utf-16').splitlines():
            view, name, axis, ax, ay, bx, by = row.split('\t')
            a, b = (number(ax), number(ay)), (number(bx), number(by))
            assert any(location(a, loop) == -1 for loop in contours[view]), 'Witness must start at an actual contour'
            assert any(location(b, loop) == -1 for loop in contours[view]), 'Witness must start at an actual contour'
            arrows = [list(map(number, r[4:6])) for r in saved if r[0] == 'ARROW' and r[2:4] == [view, name]]
            outline = next(list(map(number, r[3:7])) for r in saved if r[0] == 'VIEW' and r[2] == view)
            assert len(arrows) == 2 and all(p[0] < outline[0] or p[0] > outline[2] for p in arrows), 'Shared depth baseline must be outside the profile'
            text = next(list(map(number, r[4:6])) for r in saved if r[0] == 'TEXT' and r[2:4] == [view, name])
            if arrows[0][0] < outline[0]:
                assert text[0] < arrows[0][0] - TOLERANCE, 'Shared depth label must be left of its external baseline'
            else:
                assert text[0] > arrows[0][0] + TOLERANCE, 'Shared depth label must be right of its external baseline'
            witnesses[(view, name)] = int(axis), a, b
    failures = crossings(saved, contours, witnesses)
    assert not failures, ('Persisted dimensions cross profile material', failures[:5])
    with pdfplumber.open(next(directory.glob('PDF/*.pdf'))) as pdf:
        assert len(pdf.pages) == 1
        page = pdf.pages[0]
        black = []
        for line in page.lines:
            if line['stroking_color'] == (0, 0, 0) and line['linewidth'] >= 0.5 * 72 / 25.4:
                first, last = line['pts'][0], line['pts'][-1]
                black.append(((first[0] * PDF_UNIT, (page.height - first[1]) * PDF_UNIT),
                              (last[0] * PDF_UNIT, (page.height - last[1]) * PDF_UNIT)))
        for loops in contours.values():
            for polygon in loops:
                for point in polygon:
                    assert min(point_distance(point, first, last) for first, last in black) < 0.00002, ('Contour audit differs from PDF geometry', point)
    if before:
        original = records(external_path(before))
        boxes = {r[2]: list(map(number, r[3:7])) for r in saved if r[0] == 'VIEW'}
        old_boxes = {r[2]: list(map(number, r[3:7])) for r in original if r[0] == 'VIEW'}
        shifted = {}
        for view, loops in contours.items():
            old, new = old_boxes[view], boxes[view]
            assert math.isclose(old[2] - old[0], new[2] - new[0], abs_tol=TOLERANCE)
            assert math.isclose(old[3] - old[1], new[3] - new[1], abs_tol=TOLERANCE)
            dx, dy = (old[0] + old[2] - new[0] - new[2]) / 2, (old[1] + old[3] - new[1] - new[3]) / 2
            shifted[view] = [[(x + dx, y + dy) for x, y in polygon] for polygon in loops]
        previous = crossings(original, shifted)
        assert previous, 'The original fixture must reproduce the reported material crossings'
        print(f'Before: {len(previous)} graphical segments crossed the material; after: 0')
    print(f'PASS: {len(contours)} profile views; only {len(witnesses)} declared model-linked extension pairs may cross material; contours match black PDF geometry')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', required=True, type=Path)
    parser.add_argument('--before', type=Path)
    args = parser.parse_args()
    check(args.directory, args.before)
