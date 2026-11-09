"""Versatz einzelner Bereiche: um wie viele Pixel muss der Import verschoben werden?

    .venv/bin/python tools/pptx_rundlauf/versatz.py ref.png import.png x,y,w,h [...]
"""
import sys

from PIL import Image, ImageChops, ImageOps, ImageStat


def versatz(ref, imp, box, r=6):
    x, y, w, h = box
    a = ImageOps.grayscale(ref.crop((x, y, x + w, y + h)))
    best = None
    for dy in range(-r, r + 1):
        for dx in range(-r, r + 1):
            b = ImageOps.grayscale(imp.crop((x + dx, y + dy, x + dx + w, y + dy + h)))
            d = ImageStat.Stat(ImageChops.difference(a, b)).mean[0]
            if best is None or d < best[0]:
                best = (d, dx, dy)
    return best


if __name__ == "__main__":
    ref, imp = Image.open(sys.argv[1]).convert("RGB"), Image.open(sys.argv[2]).convert("RGB")
    for arg in sys.argv[3:]:
        box = tuple(int(v) for v in arg.split(","))
        d, dx, dy = versatz(ref, imp, box)
        print(f"{arg:22} dx {dx:+d} dy {dy:+d}  Rest {d:.2f}")
