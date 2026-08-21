from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent.parent
SCALE = 8


def quadratic(start, control, end, steps=20):
    for index in range(1, steps + 1):
        t = index / steps
        inverse = 1 - t
        yield (
            inverse * inverse * start[0] + 2 * inverse * t * control[0] + t * t * end[0],
            inverse * inverse * start[1] + 2 * inverse * t * control[1] + t * t * end[1],
        )


def make_icon(color, destination):
    # Same geometry as the SVG path, sampled at 8× resolution for antialiasing.
    # Tangent points are calculated for the same 4 px corner radius.
    points = [(6.46, 18.66), (25.86, 6.76)]
    points.extend(quadratic((25.86, 6.76), (32, 3), (32, 10.2)))
    points.append((32, 33.8))
    points.extend(quadratic((32, 33.8), (32, 41), (25.86, 37.24)))
    points.append((6.46, 25.34))
    points.extend(quadratic((6.46, 25.34), (1, 22), (6.46, 18.66)))
    canvas = Image.new("RGBA", (36 * SCALE, 44 * SCALE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas)
    draw.polygon([(round(x * SCALE), round(y * SCALE)) for x, y in points], fill=color)
    canvas.resize((36, 44), Image.Resampling.LANCZOS).save(destination, optimize=True)


make_icon("#2FA572", ROOT / "light.png")
make_icon("#106A43", ROOT / "dark.png")
