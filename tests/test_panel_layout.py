"""Validate persisted two-view panel dimensions, source witnesses, and PDF ink."""
from collections import defaultdict
from pathlib import Path
import argparse
import math
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from test_profile_layout import check_pdf_markers, segment_hits_box
from test_profile_material import cuts_material, location, number, point_distance, PDF_UNIT, TOLERANCE
import pdfplumber


def transform_point(point, transform):
    cosine, sine, tx, ty = transform
    return cosine * point[0] - sine * point[1] + tx, sine * point[0] + cosine * point[1] + ty


def fit_transform(model, paper):
    dx, dy = model[1][0] - model[0][0], model[1][1] - model[0][1]
    px, py = paper[1][0] - paper[0][0], paper[1][1] - paper[0][1]
    length_squared = dx * dx + dy * dy
    cosine = (px * dx + py * dy) / length_squared
    sine = (py * dx - px * dy) / length_squared
    return (cosine, sine, paper[0][0] - cosine * model[0][0] + sine * model[0][1],
            paper[0][1] - sine * model[0][0] - cosine * model[0][1])


def check(directory):
    directory = external_path(directory)
    saved = [line.split('\t') for line in (directory / 'saved-dimensions.tsv').read_text('utf-16').splitlines()]
    plan = [line.split('\t') for line in (directory / 'dimension-plan.tsv').read_text('utf-16').splitlines()]
    dimensions = {(row[2], row[3]): row for row in saved if row[0] == 'DIM'}
    assert len(dimensions) == len(plan) > 0, 'A planned graphical dimension was lost'
    roles = defaultdict(set)
    for row in plan:
        key = row[2], row[3]
        assert key in dimensions
        dimension = dimensions[key]
        assert math.isclose(number(dimension[5]), number(row[8]), abs_tol=0.000005), (row, dimension)
        assert dimension[8].strip() and int(dimension[9]) == 2, ('Unlabeled native markers', dimension)
        roles[row[4]].add(row[2])
    assert len(roles['outline']) == len(roles['slot']) == 1 and roles['outline'].isdisjoint(roles['slot'])

    loops = defaultdict(lambda: defaultdict(list))
    for line in (directory / 'profile-contours.tsv').read_text('utf-16').splitlines():
        job, view, loop, index, x, y = line.split('\t')
        loops[view][int(loop)].append((number(x), number(y)))
    contours = {view: [polygons[index] for index in sorted(polygons)] for view, polygons in loops.items()}
    assert len(contours) == 2
    references = defaultdict(list)
    arrows = defaultdict(list)
    for row in saved:
        if row[0] == 'REFERENCE':
            references[row[2], row[3]].append(tuple(map(number, row[4:6])))
        elif row[0] == 'ARROW':
            arrows[row[2], row[3]].append(tuple(map(number, row[4:6])))

    transforms = {}
    overview = next(iter(roles['outline']))
    axes = {}
    for axis in (0, 1):
        overall = max((row for row in plan if row[4] == 'outline' and int(row[5]) == axis), key=lambda row: number(row[8]))
        key = overall[2], overall[3]
        model = references[key][:2]
        marker = arrows[key]
        scale = (max(p[axis] for p in marker) - min(p[axis] for p in marker)) / (max(p[axis] for p in model) - min(p[axis] for p in model))
        offset = min(p[axis] for p in marker) - scale * min(p[axis] for p in model)
        axes[axis] = scale, offset
    assert math.isclose(axes[0][0], axes[1][0], rel_tol=1e-6)
    transforms[overview] = axes[0][0], 0, axes[0][1], axes[1][1]
    for line in (directory / 'profile-witnesses.tsv').read_text('utf-16').splitlines():
        view, name, axis, ax, ay, bx, by = line.split('\t')
        a, b = (number(ax), number(ay)), (number(bx), number(by))
        model = references[view, name][:2]
        candidates = [fit_transform(model, pair) for pair in ((a, b), (b, a))]
        # Native dimension references can be returned in either endpoint order.
        # Choose the mapping using every actual contour reference in this view.
        points = [point for key in references if key[0] == view for point in references[key][:2]]
        candidates = [candidate for candidate in candidates if all(
            any(location(transform_point(point, candidate), polygon) == -1 for polygon in contours[view])
            for point in points)]
        assert candidates, ('Native references do not map to the saved view', view, name)
        transform = candidates[0]
        if view in transforms:
            assert all(math.isclose(a, b, abs_tol=0.000005) for a, b in zip(transform, transforms[view]))
        else:
            transforms[view] = transform
    assert set(transforms) == set(contours)

    requests = {(row[2], row[3]): row for row in plan}
    witnesses = {}
    for key, request in requests.items():
        a, b = [transform_point(point, transforms[key[0]]) for point in references[key][:2]]
        assert all(any(location(point, loop) == -1 for loop in contours[key[0]]) for point in (a, b)), ('A dimension reference is not on the rendered contour', key, a, b)
        axis = int(request[5])
        direction = (1, 0) if axis == 0 else (0, 1)
        if axis == 2:
            length = math.dist(a, b)
            direction = (b[0] - a[0]) / length, (b[1] - a[1]) / length
        witnesses[key] = a, b, direction
    lines = []
    for row in saved:
        if row[0] != 'LINE':
            continue
        first, last = tuple(map(number, row[4:6])), tuple(map(number, row[6:8]))
        lines.append(first + last)
        if not cuts_material(first, last, contours[row[2]]):
            continue
        a, b, direction = witnesses[row[2], row[3]]
        stations = [sum((p[i] - a[i]) * direction[i] for i in (0, 1)) for p in (first, last)]
        end = sum((b[i] - a[i]) * direction[i] for i in (0, 1))
        assert abs(stations[0] - stations[1]) <= TOLERANCE and min(abs(stations[0]), abs(stations[0] - end)) <= TOLERANCE, ('A baseline or leader crosses the material', row)

    regions = {}
    for line in (directory / 'native-view-regions.tsv').read_text('utf-16').splitlines():
        row = line.split('\t')
        regions[row[1]] = list(map(number, row[2:6]))
    for row in saved:
        if row[0] == 'LINE':
            region = regions[row[2]]
            for point in (list(map(number, row[4:6])), list(map(number, row[6:8]))):
                assert region[0] - TOLERANCE <= point[0] <= region[2] + TOLERANCE, (row, region)
                assert region[1] - TOLERANCE <= point[1] <= region[3] + TOLERANCE, (row, region)
    pdfs = list(directory.glob('PDF/*.pdf'))
    assert len(pdfs) == 1
    with pdfplumber.open(pdfs[0]) as pdf:
        assert len(pdf.pages) == 1
        page = pdf.pages[0]
        markers = check_pdf_markers(page, saved)
        for char in page.chars:
            if not char['text'].strip():
                continue
            dx = (char['x1'] - char['x0']) * 0.22
            dy = (char['bottom'] - char['top']) * 0.25
            ink = [(char['x0'] + dx) * PDF_UNIT, (page.height - char['bottom'] + dy) * PDF_UNIT,
                   (char['x1'] - dx) * PDF_UNIT, (page.height - char['top'] - dy) * PDF_UNIT]
            assert not any(segment_hits_box(line, ink) for line in lines), ('A native dimension line crosses printed ink', char['text'], ink)
        black = []
        for line in page.lines:
            if line['stroking_color'] == (0, 0, 0) and line['linewidth'] >= 0.5 * 72 / 25.4:
                a, b = line['pts'][0], line['pts'][-1]
                black.append(((a[0] * PDF_UNIT, (page.height - a[1]) * PDF_UNIT),
                              (b[0] * PDF_UNIT, (page.height - b[1]) * PDF_UNIT)))
        assert black
        for polygons in contours.values():
            for polygon in polygons:
                assert all(min(point_distance(point, a, b) for a, b in black) < 0.00002 for point in polygon), 'The audited contour differs from the PDF'
    print(f'PASS: {len(dimensions)} persisted native dimensions; {markers} labeled PDF markers; isolated roles; native references on the printed contour; ink and material checks')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    check(parser.parse_args().directory)
