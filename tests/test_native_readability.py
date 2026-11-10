"""Check exported PDF ink against every persisted native dimension and marker."""
from collections import defaultdict
from pathlib import Path
import argparse
import math
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from test_profile_layout import check_pdf_markers, number
from test_planar_layout import glyph_line_hits
import pdfplumber


def check(directory, expected_sheets=None):
    directory = external_path(directory).resolve(strict=True)
    saved = [line.split('\t') for line in (directory / 'saved-dimensions.tsv').read_text('utf-16').splitlines()]
    sheets = [row[1] for row in saved if row[0] == 'SHEET']
    dimensions = [row for row in saved if row[0] == 'DIM']
    assert dimensions and len({tuple(row[1:4]) for row in dimensions}) == len(dimensions)
    assert all(row[8].strip() for row in dimensions), 'A dimension has no printed value'
    if expected_sheets is not None:
        assert len(sheets) == expected_sheets
    segments = defaultdict(list)
    for row in saved:
        if row[0] == 'LINE':
            segments[row[1]].append(list(map(number, row[4:8])))
        elif row[0] == 'ARROWGEOMETRY':
            arrow = list(map(number, row[4:]))
            if int(arrow[8]) != 1:
                continue
            magnitude = math.hypot(arrow[3], arrow[4])
            dx, dy = arrow[3] / magnitude, arrow[4] / magnitude
            length, half_width = arrow[6], arrow[7] / 2
            polygon = [(arrow[0], arrow[1]),
                       (arrow[0] + length * dx - half_width * dy, arrow[1] + length * dy + half_width * dx),
                       (arrow[0] + length * dx + half_width * dy, arrow[1] + length * dy - half_width * dx)]
            segments[row[1]].extend([*polygon[i], *polygon[(i + 1) % 3]] for i in range(3))
        elif row[0] == 'ARC':
            data = list(map(number, row[4:]))
            cx, cy = data[10:12]
            radius = math.hypot(data[4] - cx, data[5] - cy)
            first = math.atan2(data[5] - cy, data[4] - cx)
            last = math.atan2(data[8] - cy, data[7] - cx)
            positive = bool(data[16]) == (data[15] >= 0)
            sweep = (last - first) % (2 * math.pi) if positive else -((first - last) % (2 * math.pi))
            steps = max(1, math.ceil(abs(sweep) / math.radians(1)))
            points = [(cx + radius * math.cos(first + sweep * i / steps),
                       cy + radius * math.sin(first + sweep * i / steps)) for i in range(steps + 1)]
            segments[row[1]].extend([*a, *b] for a, b in zip(points, points[1:]))
    pdfs = list((directory / 'PDF').glob('*.pdf'))
    assert len(pdfs) == 1
    markers = 0
    collisions = []
    with pdfplumber.open(pdfs[0]) as document:
        assert len(document.pages) == len(sheets)
        for sheet, page in zip(sheets, document.pages):
            markers += check_pdf_markers(page, [row for row in saved if len(row) > 1 and row[1] == sheet])
            for char in page.chars:
                if not char['text'].strip():
                    continue
                assert 0 <= char['x0'] <= char['x1'] <= page.width
                assert 0 <= char['top'] <= char['bottom'] <= page.height
                hits = glyph_line_hits(char, segments[sheet], 25.4 / 72000)
                if hits:
                    collisions.append((sheet, ascii(char['text']), hits))
    assert not collisions, ('Native dimension graphics cross printed text', collisions[:10])
    print(f'PASS: {len(sheets)} sheets; {len(dimensions)} native dimensions; {markers} labeled markers; lines, arcs and arrowheads clear of PDF text')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--expected-sheets', type=int)
    args = parser.parse_args()
    check(args.directory, args.expected_sheets)
