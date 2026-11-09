"""Wo weicht der Import ab? Graues Referenzbild, Abweichungen rot (geglättet, ab 24 Stufen).

    .venv/bin/python tools/pptx_rundlauf/heatmap.py ref.png import.png ziel.png
"""
import sys

from PIL import Image, ImageChops, ImageFilter, ImageOps

ref, imp = (Image.open(p).convert("RGB") for p in sys.argv[1:3])
d = ImageChops.difference(ref.filter(ImageFilter.GaussianBlur(1)), imp.filter(ImageFilter.GaussianBlur(1)))
m = ImageOps.grayscale(d).point(lambda v: 255 if v > 24 else 0)
grau = ImageOps.grayscale(ref).point(lambda v: 160 + v // 3).convert("RGB")
rot = Image.new("RGB", ref.size, (230, 0, 0))
Image.composite(rot, grau, m).save(sys.argv[3])
