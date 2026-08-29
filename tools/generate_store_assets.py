"""Generate MSIX icon assets from the existing ScanItEasy logo."""
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "label.png"
OUTPUT = ROOT / "packaging" / "Assets"
RESAMPLING = Image.Resampling.LANCZOS

def render_icon(size: tuple[int, int], content_ratio: float = 0.8) -> Image.Image:
    source = Image.open(SOURCE).convert("RGBA")
    target_width = max(1, round(size[0] * content_ratio))
    target_height = max(1, round(size[1] * content_ratio))
    source.thumbnail((target_width, target_height), RESAMPLING)
    canvas = Image.new("RGBA", size, (0, 0, 0, 0))
    position = ((size[0] - source.width) // 2, (size[1] - source.height) // 2)
    canvas.alpha_composite(source, position)
    return canvas

def save(name: str, size: tuple[int, int], content_ratio: float = 0.8) -> None:
    render_icon(size, content_ratio).save(OUTPUT / name, optimize=True)

def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    for scale, multiplier in ((100, 1), (200, 2), (400, 4)):
        save(f"StoreLogo.scale-{scale}.png", (50 * multiplier, 50 * multiplier))
        save(f"Square44x44Logo.scale-{scale}.png", (44 * multiplier, 44 * multiplier))
        save(f"Square150x150Logo.scale-{scale}.png", (150 * multiplier, 150 * multiplier))
    save("StoreLogo.png", (50, 50))
    save("Square44x44Logo.png", (44, 44))
    save("Square150x150Logo.png", (150, 150))
    for size in (16, 24, 32, 48, 256):
        save(f"Square44x44Logo.targetsize-{size}.png", (size, size), 0.9)

if __name__ == "__main__":
    main()