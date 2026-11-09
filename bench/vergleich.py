"""pptx-to-dom gegen PowerPoint messen: Folie für Folie SSIM, Pixelabweichung, Vergleichsbilder.

    python bench/referenz.py deck.pptx bench/out/ref          # einmal, braucht PowerPoint (macOS)
    python bench/vergleich.py deck.pptx --ref bench/out/ref [--fonts ordner] [--name x]

Wandelt mit dem gebauten Paket um (node dist/cli-Pfad über convertPptx), rendert jede Folie in
Chromium (Playwright) bei 1920x1080 und vergleicht mit den Referenzbildern. --fonts: ein Ordner
mit .ttf/.otf, die per @font-face eingebunden werden (ihr Familienname wird aus der Datei
gelesen) - für Schriften, die der Browser sonst nicht hätte. Ergebnis: bench/out/<name>/.
"""
from __future__ import annotations

import argparse
import json
import pathlib
import struct
import subprocess

from PIL import Image, ImageChops, ImageOps

HERE = pathlib.Path(__file__).resolve().parent
REPO = HERE.parent
OUT = HERE / "out"


def familie(datei: pathlib.Path) -> tuple[str, int, str] | None:
    """(Familienname, Gewicht, Stil) aus der name-/OS2-Tabelle einer TrueType/OpenType-Datei."""
    b = datei.read_bytes()
    try:
        n = struct.unpack(">H", b[4:6])[0]
        t = {}
        for i in range(n):
            tag, _cs, off, _ln = struct.unpack(">4sIII", b[12 + 16 * i:28 + 16 * i])
            t[tag.decode("latin-1")] = off
        o = t["name"]
        cnt, so = struct.unpack(">HH", b[o + 2:o + 6])
        names = {}
        for i in range(cnt):
            pid, _e, _l, nid, ln, off = struct.unpack(">HHHHHH", b[o + 6 + 12 * i:o + 18 + 12 * i])
            if pid == 3 and nid in (1, 2):
                names.setdefault(nid, b[o + so + off:o + so + off + ln].decode("utf-16-be", "ignore"))
        sub = names.get(2, "Regular").lower()
        return names.get(1, datei.stem), 700 if "bold" in sub or "fett" in sub else 400, "italic" if "italic" in sub or "kursiv" in sub else "normal"
    except Exception:  # noqa: BLE001
        return None


def font_css(ordner: pathlib.Path | None) -> str:
    if not ordner:
        return ""
    regeln = []
    for f in sorted(list(ordner.glob("*.ttf")) + list(ordner.glob("*.otf"))):
        info = familie(f)
        if info:
            fam, w, st = info
            regeln.append(f"@font-face{{font-family:'{fam}';font-weight:{w};font-style:{st};src:url('{f.as_uri()}')}}")
    return "".join(regeln)


def umwandeln(pptx: pathlib.Path) -> list[str]:
    js = ("import {readFileSync} from 'node:fs';"
          f"import {{convertPptx}} from {json.dumps((REPO / 'dist' / 'index.js').as_uri())};"
          f"const r = await convertPptx(new Uint8Array(readFileSync({json.dumps(str(pptx))})));"
          "process.stdout.write(JSON.stringify(r.slides.map(s => s.html)));")
    out = subprocess.run(["node", "--input-type=module", "-e", js], capture_output=True, text=True, check=True)
    return json.loads(out.stdout)


def rendern(slides: list[str], ziel: pathlib.Path, css: str = "") -> list[pathlib.Path]:
    from playwright.sync_api import sync_playwright
    ziel.mkdir(parents=True, exist_ok=True)
    seite = ziel / "deck.html"
    seite.write_text("<!doctype html><html><head><meta charset='utf-8'><style>"
                     "html,body{margin:0;background:#fff}.slide{position:relative;width:1920px;height:1080px;overflow:hidden;"
                     f"background:#fff}}{css}</style></head><body>{''.join(slides)}</body></html>", encoding="utf-8")
    out = []
    with sync_playwright() as pw:
        br = pw.chromium.launch(args=["--allow-file-access-from-files"])
        pg = br.new_page(viewport={"width": 1920, "height": 1080})
        pg.goto(seite.as_uri(), wait_until="networkidle")
        pg.evaluate("() => Promise.all([...document.fonts].map(f => f.load().catch(() => {})))")
        for i in range(len(slides)):
            box = pg.evaluate("(i) => { const r = document.querySelectorAll('.slide')[i].getBoundingClientRect();"
                              " return {x: r.left + scrollX, y: r.top + scrollY}; }", i)
            f = ziel / f"folie_{i + 1:02d}.png"
            pg.screenshot(path=str(f), full_page=True, clip={"x": box["x"], "y": box["y"], "width": 1920, "height": 1080})
            out.append(f)
        br.close()
    return out


def _grau(im: Image.Image, w: int = 480, h: int = 270) -> list[int]:
    return list(ImageOps.grayscale(im.convert("RGB")).resize((w, h), Image.BILINEAR).tobytes())


def ssim(a: Image.Image, b: Image.Image, w: int = 480, h: int = 270, fenster: int = 8) -> float:
    """Mittlere SSIM über 8x8-Fenster auf verkleinerten Graustufenbildern (ohne numpy)."""
    pa, pb = _grau(a, w, h), _grau(b, w, h)
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    werte = []
    for y0 in range(0, h - fenster + 1, fenster):
        for x0 in range(0, w - fenster + 1, fenster):
            xs, ys = [], []
            for y in range(y0, y0 + fenster):
                o = y * w + x0
                xs += pa[o:o + fenster]
                ys += pb[o:o + fenster]
            n = len(xs)
            mx, my = sum(xs) / n, sum(ys) / n
            vx = sum((v - mx) ** 2 for v in xs) / n
            vy = sum((v - my) ** 2 for v in ys) / n
            cov = sum((xs[k] - mx) * (ys[k] - my) for k in range(n)) / n
            werte.append(((2 * mx * my + c1) * (2 * cov + c2)) / ((mx * mx + my * my + c1) * (vx + vy + c2)))
    return sum(werte) / len(werte)


def vergleichsbild(ref: Image.Image, imp: Image.Image, ziel: pathlib.Path) -> None:
    w, h = 960, 540
    r, i = ref.convert("RGB").resize((w, h)), imp.convert("RGB").resize((w, h))
    d = ImageChops.difference(r, i).convert("L").point(lambda v: 255 if v > 40 else 0)
    rot = Image.merge("RGB", (Image.new("L", (w, h), 255), ImageChops.invert(d), ImageChops.invert(d)))
    bogen = Image.new("RGB", (w * 3 + 20, h), "white")
    bogen.paste(r, (0, 0)); bogen.paste(i, (w + 10, 0)); bogen.paste(Image.blend(r, rot, 0.55), (2 * w + 20, 0))
    bogen.save(ziel)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--ref", required=True)
    ap.add_argument("--fonts")
    ap.add_argument("--name")
    a = ap.parse_args()
    pptx = pathlib.Path(a.pptx).resolve()
    ziel = OUT / (a.name or pptx.stem)
    bilder = rendern(umwandeln(pptx), ziel / "import", font_css(pathlib.Path(a.fonts).resolve() if a.fonts else None))
    summe = []
    for f in bilder:
        rf = pathlib.Path(a.ref) / f.name
        if not rf.exists():
            continue
        s = ssim(Image.open(rf), Image.open(f))
        summe.append(s)
        vergleichsbild(Image.open(rf), Image.open(f), ziel / f"vergleich_{f.stem.split('_')[1]}.png")
        print(f"{f.stem}: SSIM {s:.4f}")
    if summe:
        print(f"Mittel: SSIM {sum(summe) / len(summe):.4f}")


if __name__ == "__main__":
    main()
