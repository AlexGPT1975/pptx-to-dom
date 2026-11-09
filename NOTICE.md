# Third-party content

## ECMA-376 preset shape definitions

`data/presetShapeDefinitions.xml` - and the module `src/presets.generated.ts` built from it, which
ships in `dist/` - is the preset shape definition annex of **ECMA-376, Office Open XML File
Formats, Part 1 (Fundamentals and Markup Language Reference)**, as distributed with the
standard (this copy via the LibreOffice project, `oox/source/drawingml/customshapes/`).

Copyright © Ecma International. Ecma publishes its standards free of charge under its own
copyright licence, which allows the standard and derivative works that help implement it to be
copied, published and distributed, provided Ecma's copyright notice and licence travel with them.
The authoritative wording is in the standard itself (https://ecma-international.org/publications-and-standards/standards/ecma-376/);
check it before publishing this package.

## Runtime dependencies

- [`jszip`](https://github.com/Stuk/jszip) - MIT or GPL-3.0 (dual licence; used under MIT)
- [`@xmldom/xmldom`](https://github.com/xmldom/xmldom) - MIT

## Not included

No fonts, icons, logos or sample decks of third parties are part of this repository. Font
metrics of a few common faces (ascent/descent ratios, measured PowerPoint baseline offsets)
are included as numbers only.
