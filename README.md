# pptx-to-dom

Convert PowerPoint (`.pptx`) slides into **editable HTML that looks like PowerPoint renders
them**: every slide becomes a `<section>` of absolutely positioned elements on a 1920×1080
stage - real text, real shapes, vector icons - ready to be edited in the browser and exported
again.

Measured against PowerPoint itself (PowerPoint → PDF → PNG vs. the HTML in Chromium, SSIM,
1 = identical): **0.98–0.99** on a deck of bento tiles, KPI cards, timelines, chevrons, tables
and photos; **0.96** on a real corporate deck with slide master, layouts and placeholders.

```js
import { convertPptx, toHtmlDocument } from "pptx-to-dom";

const result = await convertPptx(await file.arrayBuffer());   // browser File/Blob, ArrayBuffer, Uint8Array
result.slides[0].html;   // '<section class="slide" ...>...</section>'
result.fontsUsed;        // ["Calibri", ...] - load these for a faithful rendering
document.body.innerHTML = toHtmlDocument(result);
```

CLI: `npx pptx-to-dom deck.pptx deck.html`

## What it reproduces

- **The inheritance chain** presentation → master → layout → slide: colour map, theme colours
  and fonts, list styles, placeholder geometry/fill/text properties, backgrounds, and the
  master's and layout's own shapes (rails, logos, lines).
- **Shape styles** (`fillRef`/`lnRef`/`effectRef`/`fontRef`) against the theme's format scheme.
- **All 187 preset shapes** from the ECMA-376 definitions with their adjust values, plus custom
  geometry. Rectangles, rounded rectangles and ellipses stay plain CSS boxes, the rest is SVG.
- **Colours** with their modifiers (lumMod/lumOff, tint/shade, alpha, satMod, …); gradients
  blended in linear light like PowerPoint; pattern fills as tiles.
- **Outlines** (width, dash, caps, arrow heads), **outer/inner shadows and glow**.
- **Pictures**: cropping, round/rounded/custom-shaped pictures, SVG icons (`asvg:svgBlip`,
  stretched to their frame as PowerPoint does).
- **Text**: insets, anchoring, autofit, bullets and numbering with hanging indents, spacing,
  kerning as PowerPoint applies it, and PowerPoint's line metrics - every line 1.2 × size ×
  spacing, the first baseline placed as measured in PowerPoint. Text is never clipped.
- **Tables** with their table style (incl. the built-in default style), **charts** (bar, column,
  line, area, pie, doughnut from the chart's cached data), **SmartArt** (its pre-rendered
  drawing), groups incl. rotation/flip, connectors, `mc:AlternateContent`.

Not reproduced: 3-D, soft edges, reflections, WordArt warps (the slide carries
`data-pptx-flags`), EMF/WMF pictures and OLE objects without a picture (a
`data-pptx-missing="x,y,w,h"` placeholder keeps their place - fill it from a rendering of
your own if you have one).

## Options

| Option | |
|---|---|
| `sectionAttributes` | extra attributes for every `<section>` |
| `fontMetrics` | `{ "my font": [ascent, descent] }` (fractions of the em, from the font's hhea table) for faces not built in |
| `powerpointBaseline` | `{ "my font": 0.955 }` - PowerPoint's first baseline at single spacing, as a fraction of the size |

## Fonts

The HTML names the fonts the deck uses; the browser must have them (installed, or loaded via
`@font-face` - see `result.fontsUsed`). Line breaks and positions match PowerPoint only with the
same fonts.

## Development

```bash
npm install
npm test            # builds (presets + tsc) and runs the unit tests
```

`bench/` measures the converter against Microsoft PowerPoint (macOS) - see `bench/README.md`.

## Licence

MIT, see `LICENSE`. The preset shape definitions are © Ecma International, see `NOTICE.md`.
