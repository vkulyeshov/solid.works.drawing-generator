"""Check planar constraint coverage and the saved drawing against its exported PDF."""
import argparse
import math
from pathlib import Path
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from test_profile_layout import check_pdf_markers, number, segment_hits_box
import pdfplumber


def rows(path):
    return [line.split('\t') for line in path.read_text(encoding='utf-16').splitlines()]


def glyph_line_hits(char, lines, unit):
    # A rotated glyph's axis-aligned PDF box contains empty corners. Test in
    # the glyph's own coordinate system instead of treating those as ink.
    a, b, c, d, e, f = char['matrix']
    advance = char['adv']
    determinant = a * d - b * c
    if abs(c) > abs(d):
        height = (char['width'] - abs(a) * advance) / abs(c)
        bottom = (char['x0'] - e - min(0, a * advance) - min(0, c * height)) / c
    else:
        height = (char['height'] - abs(b) * advance) / abs(d)
        bottom = (char['y0'] - f - min(0, b * advance) - min(0, d * height)) / d
    box = [advance * 0.22, bottom + height * 0.25, advance * 0.78, bottom + height * 0.75]
    hits = []
    for line in lines:
        local = []
        for x, y in (line[:2], line[2:]):
            x, y = x / unit - e, y / unit - f
            local.extend(((d * x - c * y) / determinant, (-b * x + a * y) / determinant))
        if segment_hits_box(local, box):
            hits.append(line)
    return hits


def check(directory):
    directory = external_path(directory)
    log = (directory / 'run.log').read_text(encoding='utf-16')
    count = int(re.findall(r'native planned dimensions=(\d+)', log)[-1])
    group_count = int(re.findall(r'independent length groups=(\d+)', log)[-1])
    planned = rows(directory / 'planar-dimensions.tsv')[-count:]
    constraints = rows(directory / 'planar-constraints.tsv')[-group_count:]
    saved = rows(directory / 'saved-dimensions.tsv')
    dimensions = {row[3]: row for row in saved if row[0] == 'DIM'}
    assert len(dimensions) == len(planned) + 1, 'A planned measurement was lost or an extra dimension was inserted'
    assert len({row[2] for row in planned}) == len(planned), 'A native dimension has multiple owners'
    grouped = {row[1]: row for row in constraints}
    lengths = [row for row in planned if row[3] == 'edge']
    assert {row[5] for row in lengths} == {row[1] for row in constraints if row[6] == 'True'}
    assert len({row[5] for row in lengths}) == len(lengths), 'A physical length group was repeated on another view'
    for row in planned:
        native = dimensions[row[2]]
        assert native[2] == row[1], ('Wrong owning view', row, native)
        native_value = number(native[5])
        if row[3] == 'angle':
            native_value = min(native_value, math.pi - native_value)
        assert math.isclose(number(row[4]), native_value, abs_tol=1e-9), ('Saved dimension geometry changed', row, native)
        assert native[8].strip(), ('Unlabeled native arrows', native)
        if row[3] == 'edge':
            assert math.isclose(number(row[4]), number(grouped[row[5]][2]), abs_tol=5e-6)
        elif row[3] == 'angle':
            assert 0 < number(row[4]) <= math.pi / 2 + 1e-6, ('Nonlocal angular representation', row)
    omitted = [row for row in constraints if row[6] == 'False']
    if omitted:
        assert len(omitted) == 2
        a, b = omitted
        determinant = number(a[3]) * number(b[4]) - number(b[3]) * number(a[4])
        assert abs(determinant) > 0.01
        x = -sum(number(row[2]) * number(row[3]) for row in constraints if row[6] == 'True')
        y = -sum(number(row[2]) * number(row[4]) for row in constraints if row[6] == 'True')
        reconstructed = ((x * number(b[4]) - y * number(b[3])) / determinant,
                         (number(a[3]) * y - number(a[4]) * x) / determinant)
        assert all(math.isclose(value, number(row[2]), abs_tol=5e-6) for value, row in zip(reconstructed, omitted))
    regions = {row[1]: list(map(number, row[2:6])) for row in rows(directory / 'native-view-regions.tsv')}
    native_lines = []
    arrow_modes = {row[3]: int(row[4]) for row in saved if row[0] == 'ARROWMODE'}
    arrow_geometry = {}
    for row in saved:
        if row[0] == 'ARROWGEOMETRY':
            arrow_geometry.setdefault(row[3], []).append(list(map(number, row[4:])))
    inside_count = outside_count = 0
    for name, arrows in arrow_geometry.items():
        if len(arrows) == 2 and int(dimensions[name][4]) in (2, 3):
            assert arrow_modes[name] in (0, 1), ('Arrow placement still relies on Smart', name)
            first, last = arrows
            dx, dy = last[0] - first[0], last[1] - first[1]
            if arrow_modes[name] == 0:
                inside_count += 1
                assert math.hypot(dx, dy) >= first[6] + last[6] + 0.00029, ('Inside heads do not fit', name)
                assert dx * first[3] + dy * first[4] > 0, ('First head is outside despite the mode', name)
                assert -dx * last[3] - dy * last[4] > 0, ('Second head is outside despite the mode', name)
            else:
                outside_count += 1
        for arrow in arrows:
            if int(arrow[8]) != 1:
                continue
            assert math.isclose(arrow[6], 0.0025, abs_tol=1e-8), ('Arrowhead longitudinal length is incorrect', name)
            assert math.isclose(arrow[7], 0.000881635, abs_tol=1e-8), ('Arrowhead transverse height is incorrect', name)
            magnitude = math.hypot(arrow[3], arrow[4])
            dx, dy = arrow[3] / magnitude, arrow[4] / magnitude
            length, half_width = arrow[6], arrow[7] / 2
            polygon = [(arrow[0], arrow[1]),
                       (arrow[0] + length * dx - half_width * dy, arrow[1] + length * dy + half_width * dx),
                       (arrow[0] + length * dx + half_width * dy, arrow[1] + length * dy - half_width * dx)]
            native_lines.extend([*polygon[i], *polygon[(i + 1) % 3]] for i in range(3))
    assert inside_count > 0 and outside_count > 0, 'Both fitting and constrained arrow cases must be covered'
    for row in saved:
        if row[0] != 'LINE':
            continue
        points = list(map(number, row[4:8]))
        region = regions[row[2]]
        for x, y in (points[:2], points[2:]):
            assert region[0] - 0.0003 <= x <= region[2] + 0.0003, row
            assert region[1] - 0.0003 <= y <= region[3] + 0.0003, row
        native_lines.append(points)
    for row in saved:
        if row[0] != 'ARC':
            continue
        data = list(map(number, row[4:]))
        cx, cy = data[10:12]
        radius = math.hypot(data[4] - cx, data[5] - cy)
        first = math.atan2(data[5] - cy, data[4] - cx)
        last = math.atan2(data[8] - cy, data[7] - cx)
        sweep = (last - first) % (2 * math.pi)
        if not data[16]:
            sweep = -((first - last) % (2 * math.pi))
        steps = max(1, math.ceil(abs(sweep) / math.radians(1)))
        points = [(cx + radius * math.cos(first + sweep * i / steps),
                   cy + radius * math.sin(first + sweep * i / steps)) for i in range(steps + 1)]
        native_lines.extend([*a, *b] for a, b in zip(points, points[1:]))
    view_notes = [row[2] for row in saved if row[0] == 'VIEWNOTE']
    russian_detail_caption = '\u041c\u0415\u0421\u0422\u041d\u042b\u0419'
    assert view_notes and all(russian_detail_caption not in text.upper() for text in view_notes), 'An automatic detail caption remained in Russian'
    pdfs = list(directory.glob('PDF/*.pdf'))
    assert len(pdfs) == 1
    with pdfplumber.open(pdfs[0]) as pdf:
        assert len(pdf.pages) == 1
        page = pdf.pages[0]
        markers = check_pdf_markers(page, saved)
        unit = 25.4 / 72000
        for char in page.chars:
            if not char['text'].strip():
                continue
            hits = glyph_line_hits(char, native_lines, unit)
            assert not hits, ('Native line crosses printed text', ascii(char['text']), hits, char.get('matrix'))
    print(f'PASS: {len(dimensions)} saved native dimensions; {len(lengths)} unique physical length groups; '
          f'{len(omitted)} lengths independently reconstructed; {markers} PDF markers matched; '
          f'{inside_count} inside / {outside_count} outside arrow pairs; native lines and heads clear of printed text')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    check(parser.parse_args().directory)
