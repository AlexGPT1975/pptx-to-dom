# Benchmark against PowerPoint

Reference = Microsoft PowerPoint itself: the deck is saved as PDF by PowerPoint (macOS, via
AppleScript - only the benchmark's own copy is opened by name, never "the active presentation")
and rasterised with `pdftoppm` at 1920×1080. The converter's HTML is rendered in Chromium; both
are compared slide by slide (SSIM, pixel difference, side-by-side pictures).

Needs: macOS with PowerPoint, `pdftoppm` (poppler), Python with `bench/requirements.txt`
(+ `playwright install chromium`), a built package (`npm run build`).

```bash
python bench/testdeck.py                                   # bench/out/testdeck.pptx (6 slides)
python bench/referenz.py bench/out/testdeck.pptx bench/out/ref
python bench/vergleich.py bench/out/testdeck.pptx --ref bench/out/ref --fonts /path/to/fonts
python bench/heatmap.py ref.png import.png heat.png        # where it differs
python bench/versatz.py ref.png import.png x,y,w,h ...     # by how many pixels
```

`testdeck.py` takes icons from `PPTX_ICONS` (pairs `name.svg`/`name.png`) and a logo from
`PPTX_LOGO`; without them it draws neutral stand-ins. Its theme names the fonts in
`PPTX_FONT_HEAD`/`PPTX_FONT_BODY`/`PPTX_FONT_MEDIUM`.

PowerPoint only sees newly installed fonts after a restart; until then it synthesises bold and
italic, which shows up as differences in bold text.
