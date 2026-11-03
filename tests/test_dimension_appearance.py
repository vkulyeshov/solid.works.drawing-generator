"""Check persisted annotation colors and actual exported PDF line weights."""
from pathlib import Path
import argparse
import math
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
import pdfplumber


def channels(color):
    if isinstance(color, (float, int)):
        return (float(color),) * 3
    return tuple(color)


def matches(color, expected):
    values = channels(color)
    return len(values) == 3 and all(abs(value - expected) < 0.002 for value in values)


def diameter_symbol_stroke(line, page):
    # SolidWorks draws the diameter glyph as an ellipse and a short diagonal.
    # Its font stroke is separate from the annotation's dimension-line weight.
    points = line['pts']
    center_x = sum(point[0] for point in points) / len(points)
    center_y = sum(point[1] for point in points) / len(points)
    height = abs(points[-1][1] - points[0][1])
    for curve in page.curves:
        width, curve_height = curve['x1'] - curve['x0'], curve['bottom'] - curve['top']
        if curve['stroke'] and not curve['fill'] and curve['stroking_color'] == line['stroking_color']:
            if abs(width - curve_height) < 0.02 and 1.2 * curve_height < height < 1.6 * curve_height:
                if abs((curve['x0'] + curve['x1']) / 2 - center_x) < 0.02 and abs((curve['top'] + curve['bottom']) / 2 - center_y) < 0.02:
                    return True
    return False


def check(directory, color):
    directory = external_path(directory)
    gray = 96 if color == 'gray' else 0
    expected = gray / 255
    rows = [row.split('\t') for row in (directory / 'saved-dimensions.tsv').read_text(encoding='utf-16').splitlines()]
    annotations = [row for row in rows if row[0] == 'APPEARANCE']
    dimensions = [row for row in rows if row[0] == 'DIM']
    assert len(annotations) == len(dimensions) > 0
    assert all(int(row[4]) == gray * 65793 and int(row[5]) & 1 for row in annotations)
    thin, thick = 0.25 * 72 / 25.4, 0.6 * 72 / 25.4
    colored_lines, contour_lines, colored_text = 0, 0, 0
    for path in directory.glob('PDF/*.pdf'):
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                colored_lines += sum(matches(line['stroking_color'], expected) and math.isclose(line['linewidth'], thin, abs_tol=0.01) for line in page.lines)
                contour_lines += sum(matches(line['stroking_color'], 0) and math.isclose(line['linewidth'], thick, abs_tol=0.01) for line in page.lines)
                colored_text += sum(matches(char['non_stroking_color'], expected) for char in page.chars)
                assert all(matches(line['stroking_color'], 0) or matches(line['stroking_color'], expected) for line in page.lines)
                if color == 'gray':
                    assert all(math.isclose(line['linewidth'], thin, abs_tol=0.01) or diameter_symbol_stroke(line, page) for line in page.lines if matches(line['stroking_color'], expected)), 'A dimension leader uses a different thickness'
    assert colored_lines > len(dimensions), ('Missing thin dimension lines', colored_lines)
    assert contour_lines > 0, 'Missing heavy black model contours'
    assert colored_text > len(dimensions), 'Missing dimension text color'
    print(f'PASS: {len(annotations)} persisted {color} annotations; {colored_lines} thin lines; {contour_lines} contour lines')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', required=True, type=Path)
    parser.add_argument('--color', choices=('gray', 'black'), default='gray')
    args = parser.parse_args()
    check(args.directory, args.color)
