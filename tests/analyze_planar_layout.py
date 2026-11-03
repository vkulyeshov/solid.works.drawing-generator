"""Independently inspect the optional native candidate audit, without opening CAD."""
import argparse
import itertools
from pathlib import Path


def cross(a, b, c, d):
    dx, dy = b[0] - a[0], b[1] - a[1]
    ex, ey = d[0] - c[0], d[1] - c[1]
    determinant = dx * ey - dy * ex
    if abs(determinant) < 1e-14:
        return False
    t = ((c[0] - a[0]) * ey - (c[1] - a[1]) * ex) / determinant
    u = ((c[0] - a[0]) * dy - (c[1] - a[1]) * dx) / determinant
    return 1e-4 < t < 1 - 1e-4 and 1e-4 < u < 1 - 1e-4


def inside(p, polygon):
    signs = []
    for a, b in zip(polygon, polygon[1:] + polygon[:1]):
        signs.append((b[0] - a[0]) * (p[1] - a[1]) - (b[1] - a[1]) * (p[0] - a[0]))
    return min(signs) > 1e-12 or max(signs) < -1e-12


def line_hits(line, polygon):
    a, b = line[:2], line[2:]
    return inside(a, polygon) or inside(b, polygon) or any(
        cross(a, b, c, d) for c, d in zip(polygon, polygon[1:] + polygon[:1]))


def conflict(a, b, crossings=True):
    for first in a['text']:
        for last in b['text']:
            if any(inside(p, last) for p in first) or any(inside(p, first) for p in last):
                return 'text overlap'
            if any(cross(p, q, r, s)
                   for p, q in zip(first, first[1:] + first[:1])
                   for r, s in zip(last, last[1:] + last[:1])):
                return 'text overlap'
    for first, last in [(a, b), (b, a)]:
        if any(line_hits(line, polygon) for line in first['line'] for polygon in last['text']):
            return 'line through text'
        if crossings and any(cross(line[:2], line[2:], other[:2], other[2:])
                             for line in first['base'] for other in last['line' if crossings is True else 'base']):
            return 'baseline crossing'
    return None


def solve(remaining, chosen, crossings):
    if not remaining:
        return chosen
    options = {name: [c for c in candidates if all(not conflict(c, other, crossings) for other in chosen.values())]
               for name, candidates in remaining.items()}
    name = min(options, key=lambda key: len(options[key]))
    for candidate in options[name]:
        result = solve({key: value for key, value in remaining.items() if key != name},
                       {**chosen, name: candidate}, crossings)
        if result is not None:
            return result
    return None


def read_audit(path, view):
    result = {}
    current = {}
    for line in path.read_text(encoding='utf-16').splitlines():
        fields = line.split('\t')
        name, number, kind = fields[:3]
        if kind == 'view':
            current[name]['view'] = fields[3]
            continue
        values = tuple(float(v.replace(',', '.')) for v in fields[3:])
        if kind == 'position':
            if int(number) == 1:
                result[name] = []
            candidate = dict(index=int(number), text=[], line=[], base=[], position=values)
            result[name].append(candidate)
            current[name] = candidate
        elif kind == 'text':
            current[name][kind].append([values[i:i+2] for i in range(0, len(values), 2)])
        else:
            current[name][kind].append(values)
    if view:
        return {name: items for name, items in result.items()
                if items[0].get('view', name.split('@', 1)[1]) == view}
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('audit', type=Path)
    parser.add_argument('--view')
    args = parser.parse_args()
    candidates = read_audit(args.audit, args.view)
    print('Candidates:', {name: len(items) for name, items in candidates.items()})
    for crossings in [True, 'dimension lines only', False]:
        assignment = solve(candidates, {}, crossings)
        print('Include baseline crossings:', crossings, 'assignment:',
              None if assignment is None else {name: c['index'] for name, c in assignment.items()})
        if assignment is None:
            for count in range(2, len(candidates) + 1):
                failures = [subset for subset in itertools.combinations(candidates, count)
                            if solve({key: candidates[key] for key in subset}, {}, crossings) is None]
                if failures:
                    print('Smallest conflicting subsets:', failures)
                    break
