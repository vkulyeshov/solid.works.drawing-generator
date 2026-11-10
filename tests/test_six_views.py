"""Check external six-view artifacts after an isolated CAD generation run."""
from collections import Counter
from pathlib import Path
import argparse
import math
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path, external_path
import pdfplumber


def number(value):
    return float(value.replace(',', '.'))


def check(directory):
    log = (directory / 'run.log').read_text('utf-16')
    rows = [line.split('\t') for line in
            (directory / 'six-view-dimensions.tsv').read_text('utf-16').splitlines()]
    kept = [row for row in rows if row[0] == 'KEEP']
    removed = [row for row in rows if row[0] == 'REMOVE']
    owners = {row[4]: row for row in kept}
    assert len(owners) == len(kept), 'A measurement has more than one owner'
    assert not any(key.startswith('unmatched|') for key in owners)
    assert len({row[1] for row in kept}) == 2
    assert 3 <= len({(row[1], row[2]) for row in kept}) <= 6
    for row in removed:
        owner = owners[row[4]]
        assert (row[6], row[7]) == (owner[3], owner[1])
        assert math.isclose(number(row[-1]), number(owner[-1]), abs_tol=1e-9)

    transforms = {}
    for line in (directory / 'six-view-transforms.tsv').read_text('utf-16').splitlines():
        row = line.split('\t')
        transforms[(row[1], row[2])] = (row[3], list(map(number, row[4:])))
    assert len(transforms) == 8
    assert sorted(Counter(sheet for sheet, _ in transforms).values()) == [4, 4]
    orthographic = [m for role, m in transforms.values() if role == 'orthographic']
    isometric = [m for role, m in transforms.values() if role == 'isometric']
    assert len(orthographic) == 6 and len(isometric) == 2
    normals = [tuple(round(m[index]) for index in (2, 5, 8)) for m in orthographic]
    assert set(normals) == {(1, 0, 0), (-1, 0, 0), (0, 1, 0),
                            (0, -1, 0), (0, 0, 1), (0, 0, -1)}
    assert max(m[12] for m in orthographic) - min(m[12] for m in orthographic) < 1e-9
    assert all(abs(isometric[0][index] + isometric[1][index]) < 1e-8
               for index in (2, 5, 8))
    assert all(transforms[(row[1], row[2])][0] == 'orthographic' for row in kept)

    saved_path = directory / 'saved-dimensions.tsv'
    if saved_path.exists():
        saved = [line.split('\t') for line in saved_path.read_text('utf-16').splitlines()]
        sheets = [row for row in saved if row[0] == 'SHEET']
        assert len(sheets) == 2 and all(row[2] == 'True' for row in sheets)
        dimensions = [row for row in saved if row[0] == 'DIM']
        actual = {(row[1], row[2], row[3]): row for row in dimensions}
        assert len(actual) == len(dimensions) == len(kept)
        for owner in kept:
            record = actual[(owner[1], owner[2], owner[3])]
            persisted = number(record[5])
            planned = number(owner[5])
            if not math.isclose(persisted, planned, abs_tol=1e-9):
                assert 'FEATURE MODEL DIAMETER value=' not in log, (owner, record)
                # Imported radial parameters retain their model radius on reopen,
                # while their drawing annotation displays a diameter.
                assert owner[4].startswith('feature|circular-size|'), (owner, record)
                radius = number(owner[4].rsplit('|', 1)[1]) * 1e-6
                assert math.isclose(persisted, radius, abs_tol=1e-9), (owner, record)
                assert math.isclose(planned, 2 * radius, abs_tol=1e-9), (owner, record)
                assert '<MOD-DIAM>' in record[8], record
                displayed_mm = [number(x) for x in re.findall(r'\d+(?:[.,]\d+)?', record[8])]
                assert any(math.isclose(x, planned * 1000, abs_tol=0.005) for x in displayed_mm), record
        regions = {}
        for line in (directory / 'native-view-regions.tsv').read_text('utf-16').splitlines():
            fields = line.split('\t')
            regions[fields[1]] = list(map(number, fields[2:6]))
        for row in saved:
            if row[0] == 'LINE':
                region = regions[row[2]]
                points = list(map(number, row[4:8]))
                for x, y in (points[:2], points[2:]):
                    assert region[0] - 0.0003 <= x <= region[2] + 0.0003, row
                    assert region[1] - 0.0003 <= y <= region[3] + 0.0003, row
        if any(key.startswith('hole-profile|') for key in owners):
            calls = [row for row in dimensions if row[7] == 'True']
            assert len(calls) == 1, calls
            assert re.search(r'\b8\s*[xX]', calls[0][8]), calls[0]
            assert 'M12x1.5' in calls[0][8] and '6H' in calls[0][8], calls[0]
            assert '\u041d\u0410\u0421\u041a\u0412\u041e\u0417\u042c' not in calls[0][8], calls[0]
            assert math.isclose(number(calls[0][5]), 0.0105, abs_tol=1e-9)
            values = lambda prefix: [number(row[5]) for key, row in owners.items() if key.startswith(prefix)]
            pitch_values = values('pattern-circle|')
            assert len(pitch_values) == 1 and math.isclose(pitch_values[0], 0.1, abs_tol=1e-9)
            assert sorted(round(x, 6) for x in values('pattern|')) == [0.028, 0.068]
            assert sorted(round(x, 6) for x in values('pattern-boundary|')) == [0.007]
            overall = next(row for row in kept if math.isclose(number(row[5]), 0.077, abs_tol=1e-9))
            assert overall[1] == sheets[0][1], overall
            profile_matrix = transforms[tuple(overall[1:3])][1]
            overall_arrows = [row for row in saved if row[0] == 'ARROW' and row[1:4] == overall[1:4]]
            assert len(overall_arrows) == 2
            view_boxes = {(row[1], row[2]): list(map(number, row[3:7])) for row in saved if row[0] == 'VIEW'}
            profile_box = view_boxes[tuple(overall[1:3])]
            vertical_profile = abs(number(overall_arrows[0][4]) - number(overall_arrows[1][4])) < 1e-9
            if vertical_profile:
                assert max(number(row[4]) for row in overall_arrows) < profile_box[0]
            else:
                assert max(number(row[5]) for row in overall_arrows) < profile_box[1]
            for size in (0.003, 0.006, 0.009):
                step = next(row for row in kept if math.isclose(number(row[5]), size, abs_tol=1e-9))
                assert step[1:3] == overall[1:3], step
                step_arrows = [row for row in saved if row[0] == 'ARROW' and row[1:4] == step[1:4]]
                assert len(step_arrows) == 2
                if vertical_profile:
                    assert min(number(row[4]) for row in step_arrows) > profile_box[2]
                else:
                    assert min(number(row[5]) for row in step_arrows) > profile_box[3]
            pitch_length = next(row for key, row in owners.items() if key.startswith('pattern|') and math.isclose(number(row[5]), 0.028, abs_tol=1e-9))
            offset = next(row for key, row in owners.items() if key.startswith('pattern-boundary|'))
            for dimension, above in ((pitch_length, True), (offset, False)):
                matrix = transforms[tuple(dimension[1:3])][1]
                arrows = [row for row in saved if row[0] == 'ARROW' and row[1:4] == dimension[1:4]]
                assert len(arrows) == 2
                box = view_boxes[tuple(dimension[1:3])]
                if abs(number(arrows[0][4]) - number(arrows[1][4])) < 1e-9:
                    if above:
                        assert max(number(row[4]) for row in arrows) < box[0]
                    else:
                        assert min(number(row[4]) for row in arrows) > box[2]
                elif above:
                    assert min(number(row[5]) for row in arrows) > box[3]
                else:
                    assert max(number(row[5]) for row in arrows) < box[1]
            flange = next(row for row in kept if math.isclose(number(row[5]), 0.006, abs_tol=1e-9))
            assert flange[1] == sheets[0][1], flange
            chamfers = [row for row in dimensions if 'Chamfer1' in row[6]]
            assert len(chamfers) == 1 and chamfers[0][4] == '3', chamfers
            assert '\u0424\u0430\u0441\u043a\u0430 1x' in chamfers[0][8] and '45' in chamfers[0][8], chamfers
            assert not any(row[4] == '2' and math.isclose(number(row[5]), 0.001, abs_tol=1e-9) for row in dimensions)
            assert sum(math.isclose(number(row[5]), 0.026, abs_tol=1e-9) for row in kept) == 1
            if 'FEATURE MODEL DIAMETER value=' in log:
                assert not any(math.isclose(number(row[5]), size, abs_tol=1e-9)
                               for row in kept for size in (0.013, 0.057))
                pitch = next(row for key, row in owners.items() if key.startswith('pattern-circle|'))
                pitch_id = (pitch[1], pitch[2], pitch[3])
                pitch_record = actual[pitch_id]
                assert pitch_record[4] == '6' and pitch_record[9] == '1', pitch_record
                assert '\u043e\u0442\u0432.' in pitch_record[8], pitch_record
                assert '\u0434\u043e \u043e\u0441\u0456' in pitch_record[8], pitch_record
                arrows = [row for row in saved if row[0] == 'ARROW' and tuple(row[1:4]) == pitch_id]
                assert len(arrows) == 1
                matrix = transforms[pitch_id[:2]][1]
                radius_on_paper = math.hypot(number(arrows[0][4]) - matrix[9], number(arrows[0][5]) - matrix[10])
                assert math.isclose(radius_on_paper, 0.05 * matrix[12], abs_tol=1e-6)
                if 'FEATURE OPEN SLOT ' in log:
                    assert number(arrows[0][4]) > matrix[9] + 0.7 * radius_on_paper, arrows
                    hole_coordinate = radius_on_paper / math.sqrt(2)
                    for sx in (-1, 1):
                        for sy in (-1, 1):
                            clearance = math.hypot(number(arrows[0][4]) - matrix[9] - sx * hole_coordinate,
                                                   number(arrows[0][5]) - matrix[10] - sy * hole_coordinate)
                            assert clearance > 0.00525 * matrix[12] + 0.0003, arrows
                    slot_width = next(row for key, row in owners.items() if key.startswith('open-slot-width|'))
                    slot_depth = next(row for key, row in owners.items() if key.startswith('open-slot-depth|'))
                    assert math.isclose(number(slot_width[5]), 0.012, abs_tol=1e-9)
                    expected_depth = 0.0234 - math.sqrt(0.0194 ** 2 - 0.006 ** 2)
                    assert math.isclose(number(slot_depth[5]), expected_depth, abs_tol=1e-9)
                    assert not any(math.isclose(number(row[5]), value, abs_tol=1e-9)
                                   for row in kept for value in (0.0194, 0.008))
                    width_id = tuple(slot_width[1:4])
                    width_matrix = transforms[width_id[:2]][1]
                    width_arrows = [row for row in saved if row[0] == 'ARROW' and tuple(row[1:4]) == width_id]
                    assert len(width_arrows) == 2
                    assert min(number(row[5]) for row in width_arrows) > width_matrix[10] + 0.062 * width_matrix[12]
                    references = [list(map(number, row[4:7])) for row in saved
                                  if row[0] == 'REFERENCE' and tuple(row[1:4]) == tuple(slot_depth[1:4])]
                    if references:
                        assert any(math.isclose(math.hypot(point[1], point[2]), 0.0194, abs_tol=1e-9)
                                   for point in references), references
                        assert any(math.isclose(point[1], 0.0234, abs_tol=1e-9) for point in references), references
                    pitch_text = next(row for row in saved if row[0] == 'TEXT' and tuple(row[1:4]) == pitch_id)
                    assert number(pitch_text[4]) > matrix[9] + 0.062 * matrix[12], pitch_text
                    for size in (0.059, 0.124):
                        diameter = next(row for row in dimensions if math.isclose(number(row[5]), size, abs_tol=1e-9))
                        label = next(row for row in saved if row[0] == 'TEXT' and row[1:4] == diameter[1:4])
                        assert number(label[4]) < matrix[9] - 0.062 * matrix[12], label
                else:
                    raise AssertionError('The Matochina keyway must use exposed circle/slot references')
                width = next(row for row in kept if math.isclose(number(row[5]), 0.098, abs_tol=1e-9))
                assert width[1] == sheets[0][1] and width[2] != pitch[2]
                assert actual[(width[1], width[2], width[3])][4] == '2'
                for suffix, value in (('|19400', 0.0388), ('|29400', 0.0588)):
                    size = next(row for key, row in owners.items() if key.startswith('feature|circular-size|') and key.endswith(suffix))
                    assert size[1] == sheets[0][1], size
                    record = actual[(size[1], size[2], size[3])]
                    assert math.isclose(number(record[5]), value, abs_tol=1e-9) and '<MOD-DIAM>' in record[8], record
                step = next(row for row in kept if math.isclose(number(row[5]), 0.003, abs_tol=1e-9))
                assert step[1] == sheets[0][1]
                assert not any(math.isclose(number(row[5]), 0.01, abs_tol=1e-9) and actual[(row[1], row[2], row[3])][4] == '2' for row in kept)

    pdfs = list(directory.rglob('*.pdf'))
    assert len(pdfs) == 1, pdfs
    with pdfplumber.open(pdfs[0]) as pdf:
        assert len(pdf.pages) == 2
        texts = []
        for page in pdf.pages:
            texts.append(page.extract_text() or '')
            assert abs(page.width * 25.4 / 72 - 210) < 0.1
            assert abs(page.height * 25.4 / 72 - 297) < 0.1
            for char in page.chars:
                assert 0 <= char['x0'] <= char['x1'] <= page.width
                assert 0 <= char['top'] <= char['bottom'] <= page.height
        if any(key.startswith('hole-profile|') for key in owners):
            assert not re.search(r'\b35[.,]36\b', '\n'.join(texts))
    assert 'all graphics inside the drawing area' in (directory / 'run.log').read_text('utf-16')
    print(f'PASS: two A4 sheets, six sides, opposite isometrics, '
          f'{len(kept)} unique measurement owners, {len(removed)} audited repetitions removed')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--directory', type=Path,
                        default=work_path('six_views/run01/final-output'))
    check(external_path(parser.parse_args().directory))
