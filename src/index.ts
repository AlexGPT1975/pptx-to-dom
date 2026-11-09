// pptx-to-dom: convert PowerPoint (.pptx) slides into editable HTML that looks like
// PowerPoint renders them. Works in browsers and Node (>= 18).
import JSZip from "jszip";
import { CANVAS_H, CANVAS_W, ConvertOptions, Pkg, PptxImportError, SlideResult, esc, fitSpace, renderSlide } from "./importer.js";
import { attr, find, findAll, kids, rId, textOf } from "./xml.js";

export { PptxImportError };
export type { ConvertOptions, SlideResult };
export { shapePaths, adjust } from "./geometry.js";

export interface ConvertResult {
  /** One <section class="slide"> per visible slide, 1920x1080 px, absolutely positioned content. */
  slides: SlideResult[];
  width: number;
  height: number;
  /** The first slide's biggest text, else the document title. */
  title: string;
  /** Every font face the slides use - load these (e.g. @font-face) for a faithful rendering. */
  fontsUsed: string[];
}

type Input = ArrayBuffer | Uint8Array | Blob;

/** Convert a whole .pptx. Throws PptxImportError for anything that is not a slide-bearing package. */
export async function convertPptx(input: Input, options: ConvertOptions = {}): Promise<ConvertResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input as ArrayBuffer);
  } catch {
    throw new PptxImportError("Datei ist kein gültiges .pptx (kein ZIP-Container).");
  }
  const files: Record<string, Uint8Array> = {};
  await Promise.all(Object.values(zip.files).filter((f) => !f.dir).map(async (f) => { files[f.name] = await f.async("uint8array"); }));
  return convertFiles(files, options);
}

/** Same as convertPptx, from the package's already unzipped parts (path -> bytes). */
export function convertFiles(files: Record<string, Uint8Array>, options: ConvertOptions = {}): ConvertResult {
  const pkg = new Pkg(files);
  const root = fitSpace(pkg.slideW, pkg.slideH);
  const presRels = pkg.relMap("ppt/presentation.xml");
  const paths = findAll(pkg.pres, "p:sldIdLst/p:sldId").map((s) => presRels[rId(s, "id") ?? ""]).filter(Boolean);
  if (!paths.length) throw new PptxImportError("Präsentation enthält keine Folien.");
  const slides: SlideResult[] = [];
  paths.forEach((path, i) => {
    try {
      if (!pkg.has(path)) throw new Error(path);
      if (["0", "false"].includes(attr(pkg.xml(path), "show") ?? "")) return; // hidden slide
      slides.push(renderSlide(pkg, path, root, i + 1, options));
    } catch {
      // one broken slide must not fail the whole import
      slides.push({ number: i + 1, flags: ["fehler"], html: `<section class="slide" data-pptx-import="1" data-pptx-flags="fehler"></section>` });
    }
  });
  return { slides, width: CANVAS_W, height: CANVAS_H, title: deckTitle(pkg, paths[0]), fontsUsed: [...pkg.faces].sort() };
}

function deckTitle(pkg: Pkg, first: string): string {
  let best = -1, text = "";
  try {
    for (const p of findAll(pkg.xml(first), ".//a:p")) {
      const runs = findAll(p, "a:r");
      const t = runs.map((r) => textOf(find(r, "a:t"))).join("").trim();
      if (!t) continue;
      const sz = Math.max(...runs.map((r) => parseInt(attr(find(r, "a:rPr"), "sz") ?? "1800", 10)));
      if (sz > best) { best = sz; text = t; }
    }
  } catch { /* fall through */ }
  if (text) return text.slice(0, 120);
  if (pkg.has("docProps/core.xml")) {
    const t = kids(pkg.xml("docProps/core.xml")).find((e) => e.localName === "title");
    if (t && textOf(t).trim()) return textOf(t).trim();
  }
  return "Präsentation";
}

/** A standalone HTML document showing all slides one below the other (for previews and tests). */
export function toHtmlDocument(result: ConvertResult, opts: { css?: string } = {}): string {
  return `<!doctype html><html lang="de"><head><meta charset="utf-8"><title>${esc(result.title)}</title><style>
body{margin:0;background:#E9ECEF}
.slide{position:relative;width:${result.width}px;height:${result.height}px;overflow:hidden;margin:0 auto 24px;background:#fff}
${opts.css ?? ""}</style></head><body>
${result.slides.map((s) => s.html).join("\n")}
</body></html>`;
}
