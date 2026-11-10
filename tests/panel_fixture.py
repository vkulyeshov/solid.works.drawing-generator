"""Read ordered planar boundary fixtures captured from copied native parts."""
from pathlib import Path
import argparse
import json
import math
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path

TOLERANCE_M = 0.000005
TOPOLOGY_TOLERANCE_M = 0.000000001
MILLIMETRES_PER_METRE = 1000


def boundary(path):
    data = json.loads(Path(path).read_text('utf-8-sig'))
    loops = [loop for loop in data['loops'] if loop['outer']]
    assert len(loops) == 1
    remaining = [segment for segment in loops[0]['segments'] if math.dist(*segment) > TOPOLOGY_TOLERANCE_M]
    a, b = remaining.pop(0)
    points = [a]
    while remaining:
        points.append(b)
        for index, (first, last) in enumerate(remaining):
            if math.dist(b, first) < TOPOLOGY_TOLERANCE_M:
                b = last
                break
            if math.dist(b, last) < TOPOLOGY_TOLERANCE_M:
                b = first
                break
        else:
            raise AssertionError('The native boundary is disconnected')
        remaining.pop(index)
    assert math.dist(b, a) < TOPOLOGY_TOLERANCE_M
    return points


def render(path, output):
    from PIL import Image, ImageDraw
    points = boundary(path)
    width, height, margin = 1800, 650, 60
    image = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(image)
    low_x, high_x = min(p[0] for p in points), max(p[0] for p in points)
    low_y, high_y = min(p[1] for p in points), max(p[1] for p in points)
    view_scale_value = min((width - 2 * margin) / (high_x - low_x),
                (height - 2 * margin) / (high_y - low_y))
    def pixel(point):
        return margin + (point[0] - low_x) * view_scale_value, height - margin - (point[1] - low_y) * view_scale_value
    closed = points + points[:1]
    draw.line([pixel(p) for p in closed], fill='black', width=2)
    for index, point in enumerate(points):
        draw.text(pixel(point), str(index), fill='gray')
    image.save(external_path(output))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('geometry', type=Path)
    parser.add_argument('--image', type=Path)
    args = parser.parse_args()
    points = boundary(args.geometry)
    for index, point in enumerate(points):
        print(index, *(round(value * MILLIMETRES_PER_METRE, 5) for value in point))
    if args.image:
        render(args.geometry, args.image)
