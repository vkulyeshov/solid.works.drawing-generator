"""Validate planned ownership against a reopened, paginated native drawing."""
from pathlib import Path
from collections import defaultdict
import argparse
import math
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from test_profile_layout import check_pdf_markers, segment_hits_box


def records(path):
    return [line.split('\t') for line in path.read_text('utf-16').splitlines() if line]


def number(value):
    return float(value.replace(',', '.'))


def validate(directory, expected_sheets=None):
    directory = external_path(directory).resolve(strict=True)
    audits = list(directory.glob('*_profile-pages.tsv'))
    assert len(audits) == 1, 'Use a completed single-part paginated test output'
    rows = records(audits[0])
    expected = int(next(row[2] for row in rows if row[0] == 'JOB'))
    planned = [row for row in rows if row[0] == 'DIM']
    assert len(planned) == expected and expected > 0, 'Incomplete planned coverage'
    assert len({row[4] for row in planned}) == expected, 'Repeated physical measurement'
    owners = {(row[2], row[3], row[7]): number(row[6]) for row in planned}
    assert len(owners) == expected, 'Two requests share one native dimension'

    reopened = records(directory / 'saved-dimensions.tsv')
    dimensions = [row for row in reopened if row[0] == 'DIM']
    assert {(row[1], row[2], row[3]) for row in dimensions} == set(owners), 'Saved ownership changed'
    for row in dimensions:
        key = tuple(row[1:4])
        assert math.isclose(number(row[5]), owners[key], abs_tol=5e-6), f'Wrong model value: {key}'
        assert row[8].strip(), f'A dimension lost its text: {key}'
        assert int(row[9]) >= 2, f'A linear dimension lost its markers: {key}'

    masks = {(row[2], row[3]): tuple(map(number, row[4:8])) for row in rows if row[0] == 'MASK'}
    transforms = {(row[1], row[2]): tuple(map(number, row[3:])) for row in reopened if row[0] == 'TRANSFORM'}
    references, arrows, arrow_geometry = defaultdict(list), defaultdict(list), defaultdict(list)
    arrow_modes = {}
    for row in reopened:
        if row[0] == 'REFERENCE':
            references[tuple(row[1:4])].append(tuple(map(number, row[4:7])))
        elif row[0] == 'ARROW':
            arrows[tuple(row[1:4])].append(tuple(map(number, row[4:6])))
        elif row[0] == 'ARROWGEOMETRY':
            arrow_geometry[tuple(row[1:4])].append(tuple(map(number, row[4:])))
        elif row[0] == 'ARROWMODE':
            arrow_modes[tuple(row[1:4])] = int(row[4])
    for key in owners:
        assert len(references[key]) >= 2, f'A dimension lost its native references: {key}'
        if math.dist(*arrows[key][:2]) >= .0053:
            assert arrow_modes[key] == 0 and all(int(item[8]) == 1 for item in arrow_geometry[key]), f'Fitting inside arrowheads were replaced by outside arrows or dots: {key}'
        if key[:2] in masks:
            matrix, box = transforms[key[:2]], masks[key[:2]]
            for point in references[key][:2]:
                projected = [sum(point[i] * matrix[i*3+j] for i in range(3)) * matrix[12] + matrix[9+j] for j in range(2)]
                assert box[0]-2e-6 <= projected[0] <= box[2]+2e-6 and box[1]-2e-6 <= projected[1] <= box[3]+2e-6, f'A dimension references hidden geometry outside its crop: {key}'

    kinds = {tuple((row[2], row[3], row[7])): row[4].split('|')[1] for row in planned if row[4].startswith('panel|')}
    for width, width_kind in kinds.items():
        if width_kind != 'notch-width':
            continue
        first, last = arrows[width][:2]
        length = math.dist(first, last)
        direction = ((last[0]-first[0])/length, (last[1]-first[1])/length)
        for gap, gap_kind in kinds.items():
            if gap_kind != 'notch-location' or gap[:2] != width[:2]:
                continue
            a, b = arrows[gap][:2]
            gap_length = math.dist(a, b)
            parallel = abs((b[0]-a[0])*direction[1] - (b[1]-a[1])*direction[0]) / gap_length
            if parallel < 1e-6:
                separation = abs((a[0]-first[0])*direction[1] - (a[1]-first[1])*direction[0])
                assert separation >= .007-2e-6, f'Widths and gaps share a dimension row: {width}, {gap}'

    sheets = {row[1] for row in reopened if row[0] == 'SHEET'}
    assert sheets, 'The native drawing has no sheets'
    if expected_sheets is not None:
        assert len(sheets) == expected_sheets, f'Expected {expected_sheets} sheets; generated {len(sheets)}'
    full = [row for row in rows if row[0] == 'VIEW' and row[4] == 'full']
    cropped = [row for row in rows if row[0] == 'VIEW' and row[4] == 'crop']
    assert len(full) == len(sheets) and {row[2] for row in full} == sheets, 'Missing complete reference view'
    assert len({int(row[6]) for row in full}) == 1 and int(full[0][6]) > 0, 'A reference view lost geometry'
    assert cropped and all(int(row[6]) > 0 for row in cropped), 'An enlarged view is empty'
    assert any(number(row[5]) >= number(full[0][5]) for row in cropped), 'All local views are smaller than the overview'

    from pypdf import PdfReader
    pdfs = list((directory / 'PDF').glob('*.pdf'))
    if not pdfs:
        pdfs = list((directory / 'PDF').glob('*.PDF'))
    assert len(pdfs) == 1, 'Expected one exported PDF'
    with pdfs[0].open('rb') as stream:
        document = PdfReader(stream)
        assert len(document.pages) == len(sheets), 'PDF sheet coverage changed'
        for page in document.pages:
            assert math.isclose(float(page.mediabox.width) * 25.4 / 72, 210, abs_tol=.1)
            assert math.isclose(float(page.mediabox.height) * 25.4 / 72, 297, abs_tol=.1)
    import pdfplumber
    sheet_order = [row[1] for row in reopened if row[0] == 'SHEET']
    marker_count = 0
    with pdfplumber.open(pdfs[0]) as document:
        for sheet, page in zip(sheet_order, document.pages):
            saved = [row for row in reopened if len(row) > 1 and row[1] == sheet]
            marker_count += check_pdf_markers(page, saved)
            lines = [list(map(number, row[4:8])) for row in saved if row[0] == 'LINE']
            for char in page.chars:
                if not char['text'].strip():
                    continue
                dx = (char['x1'] - char['x0']) * .22
                dy = (char['bottom'] - char['top']) * .25
                unit = 25.4 / 72000
                ink = [(char['x0'] + dx) * unit, (page.height - char['bottom'] + dy) * unit,
                       (char['x1'] - dx) * unit, (page.height - char['top'] - dy) * unit]
                assert not any(segment_hits_box(line, ink) for line in lines), ('A native dimension line crosses printed text', sheet, char['text'], ink)
    assert marker_count == expected * 2, 'A PDF dimension lost its arrow pair'
    print(f'Native paginated audit passed: {expected} unique dimensions, {len(sheets)} A4 sheets, {len(cropped)} detail views')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path, required=True, help='External single-part test output')
    parser.add_argument('--expected-sheets', type=int, help='Require an exact native and PDF page count')
    args = parser.parse_args()
    validate(args.directory, args.expected_sheets)
