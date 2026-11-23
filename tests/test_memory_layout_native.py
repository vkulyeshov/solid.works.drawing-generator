"""Check persisted native measurements, template bounds and reference captions."""
from pathlib import Path
import argparse
from collections import defaultdict
import math
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from pypdf import PdfReader
import pdfplumber


def rows(path):
    return [line.split('\t') for line in path.read_text('utf-16').splitlines() if line]


def number(value):
    return float(value.replace(',', '.'))


def pdf_measurement_near_graphics(page, box, value, kind):
    """Imported Toolbox DisplayData can omit digits that PDF export prints."""
    expected = math.degrees(value) if kind == 3 else value * 1000
    metres_per_point = 0.0254 / 72
    margin = 0.003
    for word in page.extract_words():
        ink = [word['x0'] * metres_per_point, (page.height - word['bottom']) * metres_per_point,
               word['x1'] * metres_per_point, (page.height - word['top']) * metres_per_point]
        if (ink[2] < box[0] - margin or ink[0] > box[2] + margin or
                ink[3] < box[1] - margin or ink[1] > box[3] + margin):
            continue
        for token in re.findall(r'[-+]?\d+(?:[.,]\d+)?', word['text']):
            if math.isclose(number(token), expected, abs_tol=0.00500001):
                return True
    return False


def check(directory):
    output = external_path(directory) / 'output'
    log = (output / 'run.log').read_text('utf-16')
    batches = re.findall(r'BATCH END drawings=(\d+) failed=(\d+)', log)
    assert batches and batches[-1] == ('1', '0'), 'Native generation did not complete'
    assert 'native Auto Arrange calls=0' in log, 'The offline engine did not complete'
    saved = rows(output / 'saved-dimensions.tsv')
    native = {(row[1], row[3]): row for row in saved if row[0] == 'DIM'}
    frames = {(row[0], row[1]): list(map(number, row[3:7]))
              for row in rows(output / 'sheet-workspace.tsv') if row[2] == 'frame'}
    audited = {}
    references = []
    for row in rows(output / 'memory-layout.tsv'):
        if not row[0].isdigit():
            continue
        assert len(row) == 10, 'A verified dimension record is incomplete'
        box = list(map(number, row[6:10]))
        frame = frames[(row[0], row[1])]
        assert box[0] >= frame[0] - 1e-8 and box[1] >= frame[1] - 1e-8, row
        assert box[2] <= frame[2] + 1e-8 and box[3] <= frame[3] + 1e-8, row
        if int(row[4]) == -1:
            references.append(row[3].rsplit('|', 1)[1])
        else:
            audited[(row[1], row[3])] = row
    assert set(audited) == set(native), 'A verified dimension was lost or gained after saving'
    for key, row in audited.items():
        assert math.isclose(number(row[5]), number(native[key][5]), abs_tol=1e-9), key
    sheets = [row for row in saved if row[0] == 'SHEET']
    assert all(row[2] == 'True' for row in sheets), 'A continuation format is hidden'
    for sheet in sheets:
        assert sum(row[0] == 'FORMATLINE' and row[1] == sheet[1] for row in saved) >= 4
    pdfs = list((output / 'PDF').glob('*.pdf'))
    assert len(pdfs) == 1, 'Expected exactly one exported drawing'
    pdf = PdfReader(pdfs[0])
    assert len(pdf.pages) == len(sheets), 'The PDF lost a sheet'
    text = '\n'.join(page.extract_text() for page in pdf.pages)
    # Do not mistake an incomplete reopened COM text run for a missing value.
    # Require the actual number near this annotation's verified paper bounds.
    with pdfplumber.open(pdfs[0]) as plotted:
        pages = {sheet[1]: page for sheet, page in zip(sheets, plotted.pages)}
        for key, row in audited.items():
            if not re.search(r'\d', native[key][8]):
                assert pdf_measurement_near_graphics(pages[key[0]], list(map(number, row[6:10])),
                                                    number(native[key][5]), int(native[key][4])), (
                    'An arrow has no printed measurement near its graphics', key)
    for label in set(references):
        assert len(re.findall(r'(?<!\w)' + re.escape(label) + r'(?!\w)', text)) >= 2, (
            'A source letter has no matching detail caption', label)
    graphics = output / 'reference-graphics.tsv'
    if references and graphics.is_file():
        leaders = defaultdict(list)
        for row in rows(graphics):
            leaders[(row[0], row[1])].append(list(map(number, row[2:6])))
        unit = 72 * 1000 / 25.4
        tolerance = 0.05 * 72 / 25.4
        with pdfplumber.open(pdfs[0]) as plotted:
            pages = {sheet[1]: page for sheet, page in zip(sheets, plotted.pages)}
            for (name, sheet), lines in leaders.items():
                for line in lines[-2:]:
                    expected = [value * unit for value in line]
                    def matches(actual):
                        ends = [actual['x0'], actual['y0'], actual['x1'], actual['y1']]
                        # pdfplumber normalizes endpoint ordering; check each axis.
                        return all(abs(a-b) <= tolerance for a,b in zip(
                            sorted(expected[0::2]) + sorted(expected[1::2]),
                            sorted(ends[0::2]) + sorted(ends[1::2])))
                    assert any(matches(item) for item in pages[sheet].lines), (
                        'A verified leader was lost in PDF export', name, line)
    assert not re.search(r'\u041c\u0415\u0421\u0422\u041d\u042b\u0419|\u041c\u0410\u0421\u0428\u0422\u0410\u0411', text), 'A native UI-language detail caption remains'
    print(f'PASS: {directory}: {len(native)} persisted measurements, {len(sheets)} sheets, '
          f'{len(set(references))} matched source letters')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('directories', nargs='+', type=Path)
    for item in parser.parse_args().directories:
        check(item)
