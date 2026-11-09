// PPTX -> HTML: every slide as a <section> of absolutely positioned, editable elements on a
// 1920x1080 stage, drawn the way PowerPoint draws it. See README.md for what is covered.
import { renderChart } from "./charts.js";
import {
  COLOR_TAGS, PRESET_COLORS, RGBA, Theme, applyMods, base64, css, fromLin, hex, linearLight, opaque, patternSvg, shadeColor,
} from "./color.js";
import { adjust, pathBBox, shapePaths } from "./geometry.js";
import { El, NS, attr, attrs, find, findAll, int, kid, kids, ln, parseXml, rId, textOf } from "./xml.js";

export const EMU_PER_PX = 6350; // 914400 EMU/in at 144 px/in (1920 px / 13.333 in)
export const CANVAS_W = 1920;
export const CANVAS_H = 1080;
const SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const EXT_MIME: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", tif: "image/tiff",
  tiff: "image/tiff", svg: "image/svg+xml", webp: "image/webp",
};

export class PptxImportError extends Error {}

export interface ConvertOptions {
  /** Extra attributes on every <section> (e.g. 'data-no-chrome="true"'). */
  sectionAttributes?: string;
  /** hhea ascent/descent (fractions of the em) of fonts the browser will use, keyed lower-case. */
  fontMetrics?: Record<string, [number, number]>;
  /** PowerPoint's first baseline at single spacing as a fraction of the size, keyed lower-case. */
  powerpointBaseline?: Record<string, number>;
}

// ----------------------------------------------------------------------- package
export class Pkg {
  files: Record<string, Uint8Array>;
  private xmlCache: Record<string, El> = {};
  private uriCache: Record<string, string | null> = {};
  slideW: number;
  slideH: number;
  pres: El;
  defaultLst: El | null;
  tableStyles: Record<string, El> = {};
  defaultTableStyle = "";
  faces = new Set<string>();

  constructor(files: Record<string, Uint8Array>) {
    this.files = files;
    if (!files["ppt/presentation.xml"]) throw new PptxImportError("Datei enthält keine PowerPoint-Präsentation (ppt/presentation.xml fehlt).");
    this.pres = this.xml("ppt/presentation.xml");
    const sz = find(this.pres, "p:sldSz");
    this.slideW = int(attr(sz, "cx"), 12192000);
    this.slideH = int(attr(sz, "cy"), 6858000);
    this.defaultLst = find(this.pres, "p:defaultTextStyle");
    if (files["ppt/tableStyles.xml"]) {
      const ts = this.xml("ppt/tableStyles.xml");
      this.defaultTableStyle = attr(ts, "def") ?? "";
      for (const st of findAll(ts, "a:tblStyle")) this.tableStyles[attr(st, "styleId") ?? ""] = st;
    }
  }

  has(path: string): boolean { return path in this.files; }

  text(path: string): string { return new TextDecoder().decode(this.files[path]); }

  xml(path: string): El {
    if (!this.xmlCache[path]) this.xmlCache[path] = parseXml(this.text(path));
    return this.xmlCache[path];
  }

  rels(part: string | null): [string, string, string][] {
    if (!part) return [];
    const i = part.lastIndexOf("/");
    const folder = i >= 0 ? part.slice(0, i) : "", name = i >= 0 ? part.slice(i + 1) : part;
    const relsPath = folder ? `${folder}/_rels/${name}.rels` : `_rels/${name}.rels`;
    if (!this.has(relsPath)) return [];
    return kids(this.xml(relsPath)).filter((r) => attr(r, "TargetMode") !== "External").map((r) => {
      const t = attr(r, "Target") ?? "";
      const path = t.startsWith("/") ? t.slice(1) : normPath(folder ? `${folder}/${t}` : t);
      return [attr(r, "Id") ?? "", attr(r, "Type") ?? "", path];
    });
  }

  relMap(part: string | null): Record<string, string> {
    const out: Record<string, string> = {};
    for (const [id, , p] of this.rels(part)) out[id] = p;
    return out;
  }

  byType(part: string | null, suffix: string): string | null {
    return this.rels(part).find(([, t]) => t.endsWith(suffix))?.[2] ?? null;
  }

  dataUri(path: string | null | undefined): string | null {
    if (!path) return null;
    if (path in this.uriCache) return this.uriCache[path];
    let uri: string | null = null;
    const ext = path.split(".").pop()!.toLowerCase();
    if (this.has(path) && EXT_MIME[ext]) {               // EMF/WMF: no browser shows them
      let data = this.files[path];
      if (ext === "svg") data = svgStretch(data);
      uri = `data:${EXT_MIME[ext]};base64,${base64(data)}`;
    }
    this.uriCache[path] = uri;
    return uri;
  }
}

export function normPath(path: string): string {
  const parts: string[] = [];
  for (const seg of path.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") parts.pop(); else parts.push(seg);
  }
  return parts.join("/");
}

/** PowerPoint stretches an SVG picture to its frame; a browser keeps its aspect ratio. */
export function svgStretch(data: Uint8Array): Uint8Array {
  const s = new TextDecoder().decode(data);
  const m = s.match(/<svg\b[^>]*>/);
  if (!m || m[0].includes("preserveAspectRatio")) return data;
  const tag = m[0];
  const self = tag.endsWith("/>");
  const fixed = tag.slice(0, self ? -2 : -1) + ' preserveAspectRatio="none"' + (self ? "/>" : ">");
  return new TextEncoder().encode(s.slice(0, m.index) + fixed + s.slice(m.index! + tag.length));
}

// ----------------------------------------------------------------------- fonts
const KNOWN_METRICS: Record<string, [number, number]> = {
  // measured from the font files (hhea ascent/descent per em)
  "sparkasse rg": [0.884, 0.213], "sparkasse": [0.884, 0.213], "sparkasse head": [0.884, 0.213],
  "sparkasse lt": [0.884, 0.213], "sparkasse medium": [0.908, 0.212], "sparkasse serif": [0.921, 0.225],
  "sparkasse rg mono": [0.95, 0.25],
  arial: [0.905, 0.212], helvetica: [0.77, 0.23], calibri: [0.952, 0.269], "calibri light": [0.952, 0.269],
  aptos: [0.939, 0.282], "segoe ui": [1.079, 0.251], verdana: [1.005, 0.21], tahoma: [1.0, 0.207],
  georgia: [0.917, 0.219], "times new roman": [0.891, 0.216], candara: [0.725, 0.275], "century gothic": [1.006, 0.22],
  cambria: [0.95, 0.222],
};
// PowerPoint's first baseline at single spacing (fraction of the size), measured: most faces ~0.955
const PPT_BASE: Record<string, number> = { "sparkasse medium": 0.97, calibri: 0.92, "calibri light": 0.92 };

// ----------------------------------------------------------------------- slide context
class Ctx {
  pkg: Pkg;
  opts: ConvertOptions;
  scale: number;
  slide: El;
  layoutPath: string | null;
  layout: El | null;
  masterPath: string | null;
  master: El | null;
  theme: Theme;
  cmap: Record<string, string>;
  txTitle: El | null;
  txBody: El | null;
  txOther: El | null;
  phLayout: Map<string, El>;
  phMaster: Map<string, El>;
  n = 0;
  slideNo = 0;
  curPart: string;
  flags = new Set<string>();

  constructor(pkg: Pkg, slidePath: string, scale: number, opts: ConvertOptions) {
    this.pkg = pkg; this.opts = opts; this.scale = scale; this.curPart = slidePath;
    this.slide = pkg.xml(slidePath);
    this.layoutPath = pkg.byType(slidePath, "/slideLayout");
    this.layout = this.layoutPath && pkg.has(this.layoutPath) ? pkg.xml(this.layoutPath) : null;
    this.masterPath = this.layoutPath ? pkg.byType(this.layoutPath, "/slideMaster") : null;
    this.master = this.masterPath && pkg.has(this.masterPath) ? pkg.xml(this.masterPath) : null;
    let themePath = this.masterPath ? pkg.byType(this.masterPath, "/theme") : null;
    if (!themePath || !pkg.has(themePath)) themePath = Object.keys(pkg.files).sort().find((n) => /^ppt\/theme\/theme\d+\.xml$/.test(n)) ?? null;
    this.theme = new Theme(themePath ? pkg.xml(themePath) : null);
    const cm = find(this.master, "p:clrMap");
    this.cmap = cm ? attrs(cm) : { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" };
    for (const root of [this.layout, this.slide]) {
      const ovr = find(root, "p:clrMapOvr/a:overrideClrMapping");
      if (ovr) this.cmap = attrs(ovr);
    }
    const tx = find(this.master, "p:txStyles");
    this.txTitle = find(tx, "p:titleStyle"); this.txBody = find(tx, "p:bodyStyle"); this.txOther = find(tx, "p:otherStyle");
    this.phLayout = phMap(this.layout); this.phMaster = phMap(this.master);
  }

  uid(p: string): string { return `${p}${this.slideNo}_${++this.n}`; }

  scheme(val: string): RGBA {
    const mapped = this.cmap[val] ?? val;
    return hex(this.theme.colors[mapped] ?? this.theme.colors[val] ?? "000000");
  }

  color(el: El | null, ph: RGBA | null = null): RGBA | null {
    if (!el) return null;
    const t = ln(el);
    let c: RGBA;
    if (t === "srgbClr") c = hex(attr(el, "val"));
    else if (t === "schemeClr") {
      const v = attr(el, "val") ?? "";
      c = v === "phClr" && ph ? ph : this.scheme(v === "phClr" ? "tx1" : v);
    } else if (t === "sysClr") c = hex(attr(el, "lastClr") ?? (attr(el, "val") === "window" ? "FFFFFF" : "000000"));
    else if (t === "prstClr") c = hex(PRESET_COLORS[attr(el, "val") ?? ""] ?? "000000");
    else if (t === "scrgbClr") c = [fromLin(int(attr(el, "r")) / 100000), fromLin(int(attr(el, "g")) / 100000), fromLin(int(attr(el, "b")) / 100000), 1];
    else return null;
    return applyMods(c, el);
  }

  colorIn(parent: El | null, ph: RGBA | null = null): RGBA | null {
    for (const c of kids(parent)) if (COLOR_TAGS.has(ln(c))) return this.color(c, ph);
    return null;
  }

  font(face: string | null | undefined, ref: string | null = null): string {
    let f = face ?? "";
    if (f.startsWith("+mj")) f = this.theme.major;
    else if (f.startsWith("+mn")) f = this.theme.minor;
    if (!f && ref) f = ref === "major" ? this.theme.major : this.theme.minor;
    return f || this.theme.minor || "Arial";
  }

  metrics(face: string): [number, number] {
    const k = face.trim().toLowerCase();
    return this.opts.fontMetrics?.[k] ?? KNOWN_METRICS[k] ?? KNOWN_METRICS.arial;
  }

  /** How far PowerPoint's first baseline sits below the one CSS computes, in px. PowerPoint:
   * lines are 1.2 x size x spacing tall, the first baseline sits (b + 0.8 x (spacing - 1)) x size
   * below the text top. CSS with the same px line-height: 0.6 x spacing x size + (asc - desc)/2 x size. */
  firstLineShift(face: string, size: number, pct: number): number {
    const k = face.trim().toLowerCase();
    const [asc, desc] = this.metrics(face);
    const b = this.opts.powerpointBaseline?.[k] ?? PPT_BASE[k] ?? 0.955;
    return (b + 0.8 * (pct - 1)) * size - (0.6 * pct + (asc - desc) / 2) * size;
  }
}

function phMap(root: El | null): Map<string, El> {
  const out = new Map<string, El>();
  const tree = find(root, "p:cSld/p:spTree");
  if (!tree) return out;
  const walk = (e: El) => {
    for (const el of kids(e)) {
      if (ln(el) === "sp" || ln(el) === "pic") {
        const ph = find(el, "p:nvSpPr/p:nvPr/p:ph") ?? find(el, "p:nvPicPr/p:nvPr/p:ph");
        if (ph) {
          const idx = attr(ph, "idx") ?? "", typ = attr(ph, "type") ?? "body";
          if (idx && !out.has(`idx:${idx}`)) out.set(`idx:${idx}`, el);
          if (!out.has(`type:${typ}`)) out.set(`type:${typ}`, el);
        }
      }
      walk(el);
    }
  };
  walk(tree);
  return out;
}

// ----------------------------------------------------------------------- fills, lines, effects
const FILL_TAGS = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);
const fillEl = (parent: El | null): El | null => kids(parent).find((c) => FILL_TAGS.has(ln(c))) ?? null;

class Fill {
  constructor(public kind: "none" | "solid" | "grad" | "img", public color: RGBA | null = null,
              public stops: [number, RGBA][] = [], public angle = 90, public radial: [number, number] | null = null,
              public uri: string | null = null, public tile = false) {}

  css(): string {
    if (this.kind === "solid") return `background:${css(this.color)}`;
    if (this.kind === "grad") {
      const st = this.stops.map(([p, c]) => `${css(c)} ${(p * 100).toFixed(1)}%`).join(", ");
      if (this.radial) return `background:radial-gradient(farthest-corner at ${this.radial[0].toFixed(0)}% ${this.radial[1].toFixed(0)}%, ${st})`;
      return `background:linear-gradient(${((this.angle + 90) % 360).toFixed(1)}deg, ${st})`;
    }
    if (this.kind === "img" && this.uri)
      return `background-image:url('${this.uri}');background-size:${this.tile ? "auto" : "100% 100%"};background-repeat:${this.tile ? "repeat" : "no-repeat"}`;
    return "";
  }

  svg(ctx: Ctx, defs: string[], w: number, h: number, shade: string | null): string {
    if (this.kind === "solid") return svgPaint(shade ? shadeColor(this.color, shade) : this.color);
    if (this.kind === "grad") {
      const gid = ctx.uid("g");
      const stops = this.stops.map(([p, c]) => `<stop offset="${p.toFixed(4)}" stop-color="${css(opaque(c))}" stop-opacity="${c[3].toFixed(3)}"/>`).join("");
      if (this.radial) {
        defs.push(`<radialGradient id="${gid}" cx="${(this.radial[0] / 100).toFixed(3)}" cy="${(this.radial[1] / 100).toFixed(3)}" r="0.75">${stops}</radialGradient>`);
      } else {
        const a = (this.angle * Math.PI) / 180;
        const half = (Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a))) / 2, cx = w / 2, cy = h / 2;
        defs.push(`<linearGradient id="${gid}" gradientUnits="userSpaceOnUse" x1="${(cx - Math.cos(a) * half).toFixed(2)}" y1="${(cy - Math.sin(a) * half).toFixed(2)}" x2="${(cx + Math.cos(a) * half).toFixed(2)}" y2="${(cy + Math.sin(a) * half).toFixed(2)}">${stops}</linearGradient>`);
      }
      return `url(#${gid})`;
    }
    if (this.kind === "img" && this.uri) {
      const pid = ctx.uid("i");
      const tw = this.tile ? 12 * ctx.scale * EMU_PER_PX : w, th = this.tile ? 12 * ctx.scale * EMU_PER_PX : h;
      defs.push(`<pattern id="${pid}" patternUnits="userSpaceOnUse" width="${tw.toFixed(2)}" height="${th.toFixed(2)}"><image href="${this.uri}" width="${tw.toFixed(2)}" height="${th.toFixed(2)}" preserveAspectRatio="none"/></pattern>`);
      return `url(#${pid})`;
    }
    return "none";
  }
}

function svgPaint(c: RGBA | null): string {
  if (!c) return "none";
  const s = css(opaque(c))!;
  return c[3] >= 0.999 ? s : `${s}" fill-opacity="${c[3].toFixed(3)}`;
}

function fillOf(ctx: Ctx, el: El | null, ph: RGBA | null = null, rels: Record<string, string> = {}, group: Fill | null = null): Fill | null {
  if (!el) return null;
  const t = ln(el);
  if (t === "noFill") return new Fill("none");
  if (t === "solidFill") { const c = ctx.colorIn(el, ph); return c ? new Fill("solid", c) : new Fill("none"); }
  if (t === "grpFill") return group ?? new Fill("none");
  if (t === "gradFill") {
    let stops: [number, RGBA][] = [];
    for (const gs of findAll(el, "a:gsLst/a:gs")) { const c = ctx.colorIn(gs, ph); if (c) stops.push([int(attr(gs, "pos")) / 100000, c]); }
    if (!stops.length) return new Fill("none");
    stops.sort((a, b) => a[0] - b[0]);
    stops = linearLight(stops);
    const path = find(el, "a:path");
    if (path) {
      const r = find(path, "a:fillToRect");
      const [l, tt, rr, b] = r ? ["l", "t", "r", "b"].map((k) => int(attr(r, k)) / 1000) : [50, 50, 50, 50];
      return new Fill("grad", null, stops, 90, [(l + 100 - rr) / 2, (tt + 100 - b) / 2]);
    }
    const lin = find(el, "a:lin");
    return new Fill("grad", null, stops, lin ? int(attr(lin, "ang")) / 60000 : 90);
  }
  if (t === "pattFill") {
    const fg = ctx.colorIn(find(el, "a:fgClr"), ph) ?? [0, 0, 0, 1];
    const bg = ctx.colorIn(find(el, "a:bgClr"), ph) ?? [255, 255, 255, 1];
    const tile = patternSvg(attr(el, "prst") ?? "pct50", fg, bg, 1.5 * ctx.scale * EMU_PER_PX);
    if (tile) return new Fill("img", null, [], 90, null, tile, true);
    return new Fill("solid", [(fg[0] + bg[0]) / 2, (fg[1] + bg[1]) / 2, (fg[2] + bg[2]) / 2, Math.max(fg[3], bg[3])]);
  }
  if (t === "blipFill") {
    const uri = blipUri(ctx, find(el, "a:blip"), rels);
    return uri ? new Fill("img", null, [], 90, null, uri, !!find(el, "a:tile")) : new Fill("none");
  }
  return null;
}

function blipUri(ctx: Ctx, blip: El | null, rels: Record<string, string>): string | null {
  if (!blip) return null;
  const svg = find(blip, `a:extLst/a:ext[@uri='${SVG_EXT}']/asvg:svgBlip`);
  if (svg) { const u = ctx.pkg.dataUri(rels[rId(svg) ?? ""]); if (u) return u; }
  return ctx.pkg.dataUri(rels[rId(blip) ?? ""]);
}

function styleRef(ctx: Ctx, style: El | null, name: string): [number, RGBA | null] {
  const ref = find(style, `a:${name}`);
  return ref ? [int(attr(ref, "idx")), ctx.colorIn(ref)] : [0, null];
}

function shapeFill(ctx: Ctx, spPr: El | null, style: El | null, rels: Record<string, string>, inherited: El | null, group: Fill | null): Fill {
  const f = fillOf(ctx, fillEl(spPr), null, rels, group);
  if (f) return f;
  const [idx, ph] = styleRef(ctx, style, "fillRef");
  if (idx) {
    const lst = idx < 1000 ? ctx.theme.fills : ctx.theme.bgfills;
    const i = idx < 1000 ? idx - 1 : idx - 1001;
    if (i >= 0 && i < lst.length) return fillOf(ctx, lst[i], ph, rels) ?? new Fill("none");
  }
  return fillOf(ctx, fillEl(inherited), null, rels, group) ?? new Fill("none");
}

interface Line { color: RGBA; width: number; dash: string; cap: string; join: string; head: El | null; tail: El | null; grad: boolean }
const DASH: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], sysDot: [1, 1], dot: [1, 3], dashDot: [4, 3, 1, 3], lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

function lineOf(ctx: Ctx, spPr: El | null, style: El | null, scale: number, inherited: El | null = null): Line | null {
  const own = find(spPr, "a:ln");
  const [idx, ph] = styleRef(ctx, style, "lnRef");
  const base = idx && idx <= ctx.theme.lines.length ? ctx.theme.lines[idx - 1] : null;
  const layers = [find(inherited, "a:ln"), base, own].filter((x): x is El => !!x);
  if (!layers.length) return null;
  let w = 9525, fill: [El, RGBA | null] | null = null, dash = "solid", cap = "flat", join = "round";
  let head: El | null = null, tail: El | null = null;
  for (const l of layers) {
    if (attr(l, "w")) w = int(attr(l, "w"));
    if (attr(l, "cap")) cap = attr(l, "cap")!;
    const f = fillEl(l);
    if (f) fill = [f, l === base ? ph : null];
    const d = find(l, "a:prstDash");
    if (d) dash = attr(d, "val") ?? "solid";
    for (const j of ["round", "bevel", "miter"]) if (find(l, `a:${j}`)) join = j;
    head = find(l, "a:headEnd") ?? head;
    tail = find(l, "a:tailEnd") ?? tail;
  }
  if (!fill) return null;
  const f = fillOf(ctx, fill[0], fill[1]);
  if (!f || f.kind === "none") return null;
  const color = f.kind === "solid" ? f.color! : f.stops[0]?.[1] ?? [0, 0, 0, 1];
  return { color, width: Math.max(0.75, w * scale), dash, cap, join, head, tail, grad: f.kind === "grad" };
}

interface Shadow { dx: number; dy: number; blur: number; color: RGBA; visible: boolean; glow: [number, RGBA] | null; inner: Shadow | null }
const shadowBox = (s: Shadow): string => {
  const parts: string[] = [];
  if (s.visible) parts.push(`${s.dx.toFixed(1)}px ${s.dy.toFixed(1)}px ${s.blur.toFixed(1)}px ${css(s.color)}`);
  if (s.glow) parts.push(`0 0 ${s.glow[0].toFixed(1)}px ${(s.glow[0] / 2).toFixed(1)}px ${css(s.glow[1])}`);
  if (s.inner) parts.push(`inset ${s.inner.dx.toFixed(1)}px ${s.inner.dy.toFixed(1)}px ${s.inner.blur.toFixed(1)}px ${css(s.inner.color)}`);
  return parts.length ? `box-shadow:${parts.join(", ")}` : "";
};
const shadowDrop = (s: Shadow): string => {
  const parts: string[] = [];
  if (s.visible) parts.push(`drop-shadow(${s.dx.toFixed(1)}px ${s.dy.toFixed(1)}px ${(s.blur / 2).toFixed(1)}px ${css(s.color)})`);
  if (s.glow) parts.push(`drop-shadow(0 0 ${(s.glow[0] / 3).toFixed(1)}px ${css(s.glow[1])})`, `drop-shadow(0 0 ${(s.glow[0] / 3).toFixed(1)}px ${css(s.glow[1])})`);
  return parts.join(" ") || "none";
};

function shadowOf(ctx: Ctx, spPr: El | null, style: El | null, scale: number): Shadow | null {
  let eff = find(spPr, "a:effectLst");
  let ph: RGBA | null = null;
  if (!eff && !find(spPr, "a:effectDag")) {
    const [idx, c] = styleRef(ctx, style, "effectRef");
    ph = c;
    if (idx && idx <= ctx.theme.effects.length) eff = find(ctx.theme.effects[idx - 1], "a:effectLst");
  }
  const one = (sh: El | null): Shadow | null => {
    if (!sh) return null;
    const c = ctx.colorIn(sh, ph);
    if (!c) return null;
    const dist = int(attr(sh, "dist")) * scale, ang = (int(attr(sh, "dir")) / 60000) * Math.PI / 180;
    // PowerPoint's blurRad is the full blur extent; it reads like CSS's blur radius
    return { dx: dist * Math.cos(ang), dy: dist * Math.sin(ang), blur: int(attr(sh, "blurRad")) * scale, color: c,
             visible: c[3] > 0.01, glow: null, inner: null };
  };
  let out = one(find(eff, "a:outerShdw"));
  const glow = find(eff, "a:glow"), inner = one(find(eff, "a:innerShdw"));
  if (!glow && !inner) return out && out.visible ? out : null;
  out = out ?? { dx: 0, dy: 0, blur: 0, color: [0, 0, 0, 0], visible: false, glow: null, inner: null };
  if (glow) { const gc = ctx.colorIn(glow, ph); if (gc) out.glow = [int(attr(glow, "rad")) * scale, gc]; }
  out.inner = inner;
  return out;
}

// ----------------------------------------------------------------------- transform spaces
class Space {
  constructor(public sx: number, public sy: number, public bx: number, public by: number, public rx = 0, public ry = 0) {}
  point(x: number, y: number): [number, number] { return [this.bx + (x - this.rx) * this.sx, this.by + (y - this.ry) * this.sy]; }
  length(w: number, h: number): [number, number] { return [w * this.sx, h * this.sy]; }
  child(off: [number, number], ext: [number, number], choff: [number, number], chext: [number, number]): Space {
    const [x0, y0] = this.point(...off), [w, h] = this.length(...ext);
    return new Space(w / (chext[0] || ext[0] || 1), h / (chext[1] || ext[1] || 1), x0, y0, choff[0], choff[1]);
  }
}

function fitSpace(w: number, h: number): Space {
  const nw = w / EMU_PER_PX, nh = h / EMU_PER_PX;
  const s = nw && nh ? Math.min(CANVAS_W / nw, CANVAS_H / nh) : 1;
  return new Space(s / EMU_PER_PX, s / EMU_PER_PX, (CANVAS_W - nw * s) / 2, (CANVAS_H - nh * s) / 2);
}

interface Xfrm { x: number; y: number; w: number; h: number; rot: number; fh: boolean; fv: boolean }
function xfrmOf(spPr: El | null): Xfrm | null {
  const x = find(spPr, "a:xfrm") ?? kid(spPr, "xfrm");
  const off = find(x, "a:off"), ext = find(x, "a:ext");
  if (!x || !off || !ext) return null;
  return { x: parseFloat(attr(off, "x") ?? "0"), y: parseFloat(attr(off, "y") ?? "0"),
           w: parseFloat(attr(ext, "cx") ?? "0"), h: parseFloat(attr(ext, "cy") ?? "0"),
           rot: int(attr(x, "rot")) / 60000, fh: ["1", "true"].includes(attr(x, "flipH") ?? ""), fv: ["1", "true"].includes(attr(x, "flipV") ?? "") };
}

// ----------------------------------------------------------------------- placeholders
const phOf = (el: El): El | null =>
  find(el, "p:nvSpPr/p:nvPr/p:ph") ?? find(el, "p:nvPicPr/p:nvPr/p:ph") ?? find(el, "p:nvGraphicFramePr/p:nvPr/p:ph");

function phChain(ctx: Ctx, ph: El | null, onLayout: boolean): El[] {
  if (!ph) return [];
  const typ = attr(ph, "type") ?? "body", idx = attr(ph, "idx") ?? "";
  const toMaster = (t: string) => ({ ctrTitle: "title", subTitle: "body", obj: "body" } as Record<string, string>)[t] ?? t;
  let mtyp = toMaster(typ);
  const out: El[] = [];
  if (!onLayout) {
    const lay = (idx ? ctx.phLayout.get(`idx:${idx}`) : undefined) ?? ctx.phLayout.get(`type:${typ}`);
    if (lay) { out.push(lay); const lph = phOf(lay); if (lph) mtyp = toMaster(attr(lph, "type") ?? "body"); }
  }
  const mas = ctx.phMaster.get(`type:${mtyp}`) ?? (idx ? ctx.phMaster.get(`idx:${idx}`) : undefined);
  if (mas) out.push(mas);
  return out;
}

function phKind(ph: El | null): "title" | "body" | "other" {
  if (!ph) return "other";
  const t = attr(ph, "type") ?? "body";
  if (t === "title" || t === "ctrTitle") return "title";
  if (["body", "subTitle", "obj", "tbl", "chart", "dgm", "media", "clipArt", "pic"].includes(t)) return "body";
  return "other";
}

// ----------------------------------------------------------------------- text
type Props = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const RPR_ATTRS = ["sz", "b", "i", "u", "strike", "spc", "cap", "baseline", "kern"];

function rprMerge(base: Props, rPr: El | null): Props {
  if (!rPr) return base;
  const out: Props = { ...base };
  for (const k of RPR_ATTRS) if (attr(rPr, k) !== null) out[k] = attr(rPr, k);
  for (const c of kids(rPr)) {
    const t = ln(c);
    if (t === "solidFill" || t === "gradFill" || t === "noFill") { out.fill = c; delete out.color; }
    else if (t === "latin") out.latin = attr(c, "typeface") ?? "";
    else if (t === "highlight") out.highlight = c;
  }
  return out;
}

function pprMerge(base: Props, pPr: El | null): Props {
  if (!pPr) return base;
  const out: Props = { ...base };
  for (const k of ["marL", "indent", "algn", "rtl"]) if (attr(pPr, k) !== null) out[k] = attr(pPr, k);
  for (const c of kids(pPr)) {
    const t = ln(c);
    if (t === "lnSpc" || t === "spcBef" || t === "spcAft") {
      const pct = find(c, "a:spcPct"), pts = find(c, "a:spcPts");
      if (pct) out[t] = ["pct", int(attr(pct, "val"), 100000) / 100000];
      else if (pts) out[t] = ["pts", int(attr(pts, "val")) / 100];
    } else if (["buNone", "buChar", "buAutoNum", "buBlip"].includes(t)) out.bu = c;
    else if (t === "buClr" || t === "buClrTx") out.buClr = t === "buClr" ? c : null;
    else if (["buSzPct", "buSzPts", "buSzTx"].includes(t)) out.buSz = t === "buSzTx" ? null : c;
    else if (t === "buFont" || t === "buFontTx") out.buFont = t === "buFont" ? attr(c, "typeface") : null;
    else if (t === "defRPr") out.rpr = rprMerge(out.rpr ?? {}, c);
  }
  return out;
}

function lstLevels(lst: El | null): Record<number, El> {
  const out: Record<number, El> = {};
  for (const c of kids(lst)) { const m = ln(c).match(/^lvl(\d)pPr$/); if (m) out[+m[1]] = c; }
  return out;
}

class TextBase {
  fontrefAt = 0;
  constructor(public levels: Record<number, El>[], public fontColor: RGBA | null, public fontRef: string | null) {}
  level(lvl: number): Props {
    let props: Props = { rpr: {} };
    this.levels.forEach((lv, i) => { if (i === this.fontrefAt) props = this.applyRef(props); props = pprMerge(props, lv[lvl] ?? null); });
    if (this.fontrefAt >= this.levels.length) props = this.applyRef(props);
    return props;
  }
  private applyRef(props: Props): Props {
    // the style's colour replaces any inherited fill; later levels and runs may set their own
    if (!this.fontColor) return props;
    const rpr = { ...props.rpr };
    delete rpr.fill;
    rpr.color = this.fontColor;
    return { ...props, rpr };
  }
}

function textBase(ctx: Ctx, el: El, chain: El[], style: El | null): TextBase {
  const kind = phKind(phOf(el));
  const levels: Record<number, El>[] = kind === "other"
    ? [lstLevels(ctx.pkg.defaultLst), lstLevels(ctx.txOther)]
    : [lstLevels(kind === "title" ? ctx.txTitle : ctx.txBody)];
  for (const anc of [...chain].reverse()) levels.push(lstLevels(find(kid(anc, "txBody"), "a:lstStyle")));
  const fref = find(style, "a:fontRef");
  const tb = new TextBase(levels, fref ? ctx.colorIn(fref) : null, fref ? attr(fref, "idx") : null);
  tb.fontrefAt = levels.length;
  levels.push(lstLevels(find(kid(el, "txBody"), "a:lstStyle")));
  return tb;
}

interface Run { size: number; face: string; bold: boolean; italic: boolean; u: boolean; strike: boolean; spc: number;
                cap: string | null; baseline: number; kern: boolean; color: RGBA; nofill: boolean; highlight: RGBA | null }

function runProps(ctx: Ctx, rpr: Props, fontScale: number, ref: string | null): Run {
  const sz = (int(rpr.sz, 1800) / 100) * fontScale;
  let color: RGBA | null = null;
  const f: El | undefined = rpr.fill;
  if (f && ln(f) === "solidFill") color = ctx.colorIn(f);
  else if (f && ln(f) === "gradFill") color = ctx.colorIn(find(f, "a:gsLst/a:gs"));
  color = color ?? rpr.color ?? null;
  const face = ctx.font(rpr.latin, ref);
  ctx.pkg.faces.add(face);
  return {
    size: sz * 2 * ctx.scale * EMU_PER_PX, face, bold: ["1", "true"].includes(rpr.b), italic: ["1", "true"].includes(rpr.i),
    u: rpr.u !== undefined && rpr.u !== "none", strike: rpr.strike !== undefined && rpr.strike !== "noStrike",
    spc: int(rpr.spc) / 100, cap: rpr.cap ?? null, baseline: int(rpr.baseline),
    kern: rpr.kern !== undefined && int(rpr.kern) <= sz * 100,      // PowerPoint kerns from "kern" size on
    color: color ?? [0, 0, 0, 1], nofill: !!f && ln(f) === "noFill",
    highlight: rpr.highlight ? ctx.colorIn(rpr.highlight) : null,
  };
}

export const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escAttr = (s: string) => s.replace(/"/g, "&quot;");
// SINGLE quotes: this lands inside style="…", double quotes would close the attribute
export const faceCss = (face: string) => `font-family:'${escAttr(face.replace(/\\/g, "\\\\").replace(/'/g, "\\'"))}',Arial,sans-serif`;

function runCss(p: Run, ctx: Ctx, sh: Shadow | null = null): string {
  const out = [`font-size:${p.size.toFixed(1)}px`];
  if (!p.kern) out.push("font-kerning:none");
  out.push("font-variant-ligatures:none");
  if (p.bold) out.push("font-weight:700");
  if (p.italic) out.push("font-style:italic");
  const deco = [p.u ? "underline" : "", p.strike ? "line-through" : ""].filter(Boolean);
  if (deco.length) out.push(`text-decoration:${deco.join(" ")}`);
  if (p.spc) out.push(`letter-spacing:${(p.spc * 2 * ctx.scale * EMU_PER_PX).toFixed(2)}px`);
  if (p.cap === "all") out.push("text-transform:uppercase"); else if (p.cap === "small") out.push("font-variant:small-caps");
  if (p.baseline) out.push(`vertical-align:${(p.baseline / 1000).toFixed(0)}%;font-size:${(p.size * 0.7).toFixed(1)}px`);
  out.push(p.nofill ? "color:transparent" : `color:${css(p.color)}`);
  if (p.highlight) out.push(`background:${css(p.highlight)}`);
  if (sh && sh.visible) out.push(`text-shadow:${sh.dx.toFixed(1)}px ${sh.dy.toFixed(1)}px ${(sh.blur / 2).toFixed(1)}px ${css(sh.color)}`);
  out.push(faceCss(p.face));
  return out.join(";");
}

const spacingPx = (spec: [string, number] | undefined, size: number, ctx: Ctx): number =>
  !spec ? 0 : spec[0] === "pct" ? spec[1] * 1.2 * size : spec[1] * 2 * ctx.scale * EMU_PER_PX;

const AUTONUM: Record<string, [string, string]> = {
  arabicPeriod: ["1", "{}."], arabicParenR: ["1", "{})"], arabicParenBoth: ["1", "({})"], arabicPlain: ["1", "{}"],
  alphaLcPeriod: ["a", "{}."], alphaUcPeriod: ["A", "{}."], alphaLcParenR: ["a", "{})"], alphaUcParenR: ["A", "{})"],
  romanLcPeriod: ["i", "{}."], romanUcPeriod: ["I", "{}."], romanLcParenR: ["i", "{})"],
};
function roman(n: number): string {
  const v: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let out = "";
  for (const [k, s] of v) while (n >= k) { out += s; n -= k; }
  return out;
}
function autonum(typ: string, n: number): string {
  const [kind, fmt] = AUTONUM[typ] ?? ["1", "{}."];
  let s: string;
  if (kind === "1") s = String(n);
  else if (kind === "a" || kind === "A") { s = ""; let k = n; while (k > 0) { const r = (k - 1) % 26; s = String.fromCharCode(97 + r) + s; k = Math.floor((k - 1) / 26); } if (kind === "A") s = s.toUpperCase(); }
  else { s = roman(n); if (kind === "I") s = s.toUpperCase(); }
  return fmt.replace("{}", s);
}

function textHtml(ctx: Ctx, txBody: El, base: TextBase, box: [number, number, number, number], chain: El[],
                  sh: Shadow | null = null, origin: [number, number] = [0, 0], rotate: [number, number, number] | null = null): string {
  const paras = findAll(txBody, "a:p");
  const hasContent = paras.some((p) => findAll(p, ".//a:t").some((t) => textOf(t).trim()) || find(p, "a:fld"));
  if (!hasContent) return "";
  const bp: Record<string, string> = {};
  for (const anc of [...chain].reverse()) Object.assign(bp, attrs(find(kid(anc, "txBody"), "a:bodyPr")));
  const own = find(txBody, "a:bodyPr");
  Object.assign(bp, attrs(own));
  const autofit = find(own, "a:normAutofit");
  const s = ctx.scale;
  const lI = int(bp.lIns, 91440) * s, tI = int(bp.tIns, 45720) * s, rI = int(bp.rIns, 91440) * s, bI = int(bp.bIns, 45720) * s;
  const fontScale = autofit ? int(attr(autofit, "fontScale"), 100000) / 100000 : 1;
  const lnRed = autofit ? int(attr(autofit, "lnSpcReduction"), 0) / 100000 : 0;
  const anchor = ({ t: "flex-start", ctr: "center", b: "flex-end", just: "center", dist: "center" } as Record<string, string>)[bp.anchor ?? "t"] ?? "flex-start";
  const nowrap = bp.wrap === "none";
  const vert = bp.vert ?? "horz";
  const counters: Record<number, number> = {};
  const html: string[] = [];
  let firstShift: number | null = null;
  paras.forEach((p, i) => {
    const pPr = find(p, "a:pPr");
    const lvl = pPr ? int(attr(pPr, "lvl")) + 1 : 1;
    const props = pprMerge(base.level(lvl), pPr);
    const runs = kids(p).filter((c) => ["r", "br", "fld"].includes(ln(c)));
    const rp = runs.map((r) => [r, runProps(ctx, rprMerge(props.rpr ?? {}, find(r, "a:rPr")), fontScale, base.fontRef)] as [El, Run]);
    const endP = runProps(ctx, rprMerge(props.rpr ?? {}, find(p, "a:endParaRPr")), fontScale, base.fontRef);
    const sizes = rp.filter(([r]) => ln(r) !== "br").map(([, q]) => q.size);
    const size = sizes.length ? Math.max(...sizes) : endP.size;
    const lnSp = props.lnSpc as [string, number] | undefined;
    let lineH: number, pct: number;
    if (lnSp && lnSp[0] === "pts") { lineH = lnSp[1] * 2 * s * EMU_PER_PX; pct = size ? lineH / (1.2 * size) : 1; }
    else { pct = Math.max(0.1, (lnSp ? lnSp[1] : 1) - lnRed); lineH = 1.2 * size * pct; }
    const first = rp[0]?.[1] ?? endP;
    // a line is 1.2 x size x spacing tall whatever faces it mixes: the paragraph's own strut
    // sets the line, the runs add nothing (line-height:0)
    const pcss = [`line-height:${lineH.toFixed(2)}px`, `font-size:${size.toFixed(1)}px`, "margin:0", faceCss(first.face)];
    const algn = ({ ctr: "center", r: "right", just: "justify", dist: "justify" } as Record<string, string>)[props.algn ?? "l"] ?? "left";
    if (algn !== "left") pcss.push(`text-align:${algn}`);
    const marL = int(props.marL) * s, indent = int(props.indent) * s;
    if (i > 0) { const sb = spacingPx(props.spcBef, size, ctx); if (sb) pcss.push(`margin-top:${sb.toFixed(1)}px`); }
    const sa = spacingPx(props.spcAft, size, ctx);
    if (sa && i < paras.length - 1) pcss.push(`margin-bottom:${sa.toFixed(1)}px`);
    if (firstShift === null) firstShift = ctx.firstLineShift(first.face, size, pct);
    let bullet = "";
    const bu: El | null = props.bu ?? null;
    const hasText = rp.some(([r]) => ln(r) !== "br" && textOf(find(r, "a:t")));
    if (bu && ln(bu) !== "buNone" && hasText) {
      let ch: string;
      if (ln(bu) === "buAutoNum") {
        const n = (counters[lvl] ?? int(attr(bu, "startAt"), 1) - 1) + 1;
        counters[lvl] = n;
        ch = autonum(attr(bu, "type") ?? "arabicPeriod", n);
      } else ch = ln(bu) === "buChar" ? attr(bu, "char") ?? "•" : "•";
      const bc = props.buClr ? ctx.colorIn(props.buClr) : first.color;
      let bsz = first.size;
      const bs: El | null = props.buSz ?? null;
      if (bs && ln(bs) === "buSzPct") bsz *= int(attr(bs, "val"), 100000) / 100000;
      else if (bs && ln(bs) === "buSzPts") bsz = (int(attr(bs, "val")) / 100) * 2 * s * EMU_PER_PX;
      let bface: string = props.buFont || first.face;
      if (bface === "Wingdings" || bface === "Symbol") { ch = ({ "§": "■", "Ø": "➢", "ü": "✓" } as Record<string, string>)[ch] ?? "•"; bface = "Arial"; }
      bullet = `<span class="bul" style="display:inline-block;line-height:0;width:${Math.max(-indent, 0).toFixed(1)}px;text-indent:0;font-size:${bsz.toFixed(1)}px;color:${css(bc)};${faceCss(bface)}">${esc(ch)}</span>`;
      for (const k of Object.keys(counters)) if (+k > lvl) delete counters[+k];
    } else if (hasText) delete counters[lvl];
    if (marL || indent) { pcss.push(`padding-left:${marL.toFixed(1)}px`); if (indent) pcss.push(`text-indent:${indent.toFixed(1)}px`); }
    pcss.push(nowrap ? "white-space:pre" : "white-space:pre-wrap;overflow-wrap:break-word");
    const spans: string[] = [];
    for (const [r, q] of rp) {
      if (ln(r) === "br") { spans.push("<br>"); continue; }
      let txt = textOf(find(r, "a:t"));
      if (ln(r) === "fld" && (attr(r, "type") ?? "").startsWith("slidenum") && ctx.slideNo) txt = String(ctx.slideNo);
      txt = txt.replace(/​/g, "").replace(/\t/g, "    ");
      if (txt) spans.push(`<span style="line-height:0;${runCss(q, ctx, sh)}">${esc(txt)}</span>`);
    }
    html.push(`<div style="${pcss.join(";")}">${bullet}${spans.join("") || "&#8203;"}</div>`);
  });
  const [l, t, r, b] = box;
  const x = l + lI - origin[0], y = t + tI - origin[1];
  const w = Math.max(0, r - l - lI - rI), h = Math.max(0, b - t - tI - bI);
  // PowerPoint's baseline offset moves every line alike, whatever the anchor: shift the box itself
  const shift = firstShift ?? 0;
  const cont = [`position:absolute;left:${x.toFixed(1)}px;top:${(y + shift).toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px`,
    "display:flex;flex-direction:column", `justify-content:${anchor}`, "box-sizing:border-box", "overflow:visible"];
  if (rotate) {
    const [deg, cx, cy] = rotate;
    cont.push(`transform:rotate(${deg.toFixed(2)}deg);transform-origin:${(cx - origin[0] - x).toFixed(1)}px ${(cy - origin[1] - y - shift).toFixed(1)}px`);
  }
  if (nowrap) {
    const a0 = attr(find(paras[0], "a:pPr"), "algn") ?? "l";
    cont.push(`align-items:${({ ctr: "center", r: "flex-end" } as Record<string, string>)[a0] ?? "flex-start"}`);
  }
  if (["vert", "eaVert", "wordArtVert"].includes(vert)) cont.push("writing-mode:vertical-rl");
  else if (vert === "vert270") cont.push("writing-mode:vertical-rl;transform:rotate(180deg)");
  const cols = int(bp.numCol, 1);
  if (cols > 1) cont.push(`column-count:${cols};column-gap:${(int(bp.spcCol) * s).toFixed(1)}px;display:block`);
  return `<div class="pp-text" style="${cont.join(";")}">${html.join("")}</div>`;
}

// ----------------------------------------------------------------------- shapes
const CSS_GEOMS = new Set(["rect", "roundRect", "ellipse"]);

function geomEl(spPr: El | null, chain: El[]): El | null {
  for (const src of [spPr, ...chain.map((c) => kid(c, "spPr"))]) {
    const g = find(src, "a:prstGeom") ?? find(src, "a:custGeom");
    if (g) return g;
  }
  return null;
}

function marker(ctx: Ctx, defs: string[], end: El | null, color: RGBA, start: boolean): string {
  if (!end || (attr(end, "type") ?? "none") === "none") return "";
  const typ = attr(end, "type");
  const size: Record<string, number> = { sm: 2, med: 3, lg: 5 };
  const w = size[attr(end, "w") ?? "med"] ?? 3, l = size[attr(end, "len") ?? "med"] ?? 3;
  const id = ctx.uid("m"), c = css(opaque(color));
  let shape: string, ref: number;
  if (typ === "oval") { shape = `<ellipse cx="${l / 2}" cy="${w / 2}" rx="${l / 2}" ry="${w / 2}" fill="${c}"/>`; ref = l / 2; }
  else if (typ === "diamond") { shape = `<path d="M0 ${w / 2} L${l / 2} 0 L${l} ${w / 2} L${l / 2} ${w} Z" fill="${c}"/>`; ref = l / 2; }
  else if (typ === "arrow") { shape = `<path d="M0 0 L${l} ${w / 2} L0 ${w}" fill="none" stroke="${c}" stroke-width="1"/>`; ref = l - 0.5; }
  else if (typ === "stealth") { shape = `<path d="M0 0 L${l} ${w / 2} L0 ${w} L${l * 0.3} ${w / 2} Z" fill="${c}"/>`; ref = l; }
  else { shape = `<path d="M0 0 L${l} ${w / 2} L0 ${w} Z" fill="${c}"/>`; ref = l; }
  defs.push(`<marker id="${id}" viewBox="0 0 ${l} ${w}" refX="${ref}" refY="${w / 2}" markerWidth="${l}" markerHeight="${w}" markerUnits="strokeWidth" orient="${start ? "auto-start-reverse" : "auto"}" overflow="visible">${shape}</marker>`);
  return ` marker-${start ? "start" : "end"}="url(#${id})"`;
}

function strokeAttrs(line: Line | null): string {
  if (!line) return ' stroke="none"';
  let out = ` stroke="${css(opaque(line.color))}" stroke-width="${line.width.toFixed(2)}"`;
  if (line.color[3] < 0.999) out += ` stroke-opacity="${line.color[3].toFixed(3)}"`;
  if (DASH[line.dash]) out += ` stroke-dasharray="${DASH[line.dash].map((v) => (v * line.width).toFixed(2)).join(" ")}"`;
  out += ` stroke-linecap="${({ rnd: "round", sq: "square" } as Record<string, string>)[line.cap] ?? "butt"}" stroke-linejoin="${line.join}"`;
  return out;
}

function flagEffects(ctx: Ctx, el: El, spPr: El | null): void {
  const eff = find(spPr, "a:effectLst");
  for (const t of ["softEdge", "reflection", "prstShdw"]) if (find(eff, `a:${t}`)) ctx.flags.add("effekt");
  if (find(spPr, "a:scene3d") || find(spPr, "a:sp3d")) ctx.flags.add("3d");
  const warp = find(kid(el, "txBody"), "a:bodyPr/a:prstTxWarp");
  if (warp && (attr(warp, "prst") ?? "textNoShape") !== "textNoShape") ctx.flags.add("wordart");
}

function renderSp(ctx: Ctx, el: El, space: Space, rels: Record<string, string>, groupFill: Fill | null, onLayout: boolean, connector: boolean): string {
  const spPr = kid(el, "spPr"), style = kid(el, "style");
  const chain = phChain(ctx, phOf(el), onLayout);
  flagEffects(ctx, el, spPr);
  let xf = xfrmOf(spPr);
  for (const anc of chain) { if (xf) break; xf = xfrmOf(kid(anc, "spPr")); }
  if (!xf) return "";
  const [x, y] = space.point(xf.x, xf.y);
  const [w, h] = space.length(xf.w, xf.h);
  if (w < 0 || h < 0) return "";
  const scale = (Math.abs(space.sx) + Math.abs(space.sy)) / 2;
  const inhFill = chain.map((c) => kid(c, "spPr")).find((sp) => fillEl(sp)) ?? null;
  const fill = connector ? new Fill("none") : shapeFill(ctx, spPr, style, rels, inhFill, groupFill);
  const line = lineOf(ctx, spPr, style, scale, chain.map((c) => kid(c, "spPr")).find((sp) => find(sp, "a:ln")) ?? null);
  const sh = shadowOf(ctx, spPr, style, ctx.scale);
  const geom = geomEl(spPr, chain);
  const prst = geom ? (ln(geom) === "prstGeom" ? attr(geom, "prst") ?? "rect" : "cust") : "rect";
  const txBody = kid(el, "txBody");
  const base = txBody ? textBase(ctx, el, chain, style) : null;
  const rot = xf.rot;
  const noPaint = fill.kind === "none" && !line;
  const textShadow = noPaint && sh ? sh : null;

  if (CSS_GEOMS.has(prst) && !(line && DASH[line.dash]) && !connector && !(line && line.grad)) {
    const lw = line ? line.width : 0;
    const st = [`position:absolute;left:${(x - lw / 2).toFixed(1)}px;top:${(y - lw / 2).toFixed(1)}px;width:${(w + lw).toFixed(1)}px;height:${(h + lw).toFixed(1)}px`, "box-sizing:border-box"];
    if (rot) st.push(`transform:rotate(${rot.toFixed(2)}deg)`);
    if (prst === "ellipse") st.push("border-radius:50%");
    else if (prst === "roundRect") {
      const radius = (Math.max(0, Math.min(50000, adjust(geom, "adj", 16667))) / 100000) * Math.min(w, h);
      if (radius > 0) st.push(`border-radius:${(radius + lw / 2).toFixed(1)}px`);
    }
    const fc = fill.css();
    if (fc) st.push(fc);
    if (line) st.push(`border:${lw.toFixed(2)}px solid ${css(line.color)}`);
    if (sh && !noPaint) { const b = shadowBox(sh); if (b) st.push(b); }
    let inner = "";
    if (txBody && base) {
      // the preset's text rectangle: a rounded rectangle insets its text by the corner
      const g = prst !== "rect" && geom ? shapePaths(geom, Math.max(w, 0.01), Math.max(h, 0.01)) : null;
      const box = g ? g.text : [0, 0, w, h] as [number, number, number, number];
      if (rot) {
        // the PPTX export of the host app reads a rotation off the element itself: the text
        // becomes the shape's sibling, turned about the same centre
        return `<div style="${st.join(";")};"></div>` + textHtml(ctx, txBody, base, box, chain, textShadow, [-x, -y], [rot, w / 2, h / 2]);
      }
      inner = textHtml(ctx, txBody, base, box, chain, textShadow, [lw / 2, lw / 2]);
    }
    return `<div style="${st.join(";")};">${inner}</div>`;
  }

  let g = geom ? shapePaths(geom, Math.max(w, 0.01), Math.max(h, 0.01)) : null;
  if (!g && geom) ctx.flags.add("form");
  if (!g) g = { paths: [{ d: `M0 0 L${w} 0 L${w} ${h} L0 ${h} Z`, fill: true, stroke: true, shade: null }], text: [0, 0, w, h] };
  const defs: string[] = [];
  const paths = g.paths.map((p) => {
    const fa = p.fill && fill.kind !== "none" ? fill.svg(ctx, defs, w, h, p.shade) : "none";
    const sa = strokeAttrs(p.stroke ? line : null);
    const mk = line && p.stroke && !p.d.trim().endsWith("Z")
      ? marker(ctx, defs, line.head, line.color, true) + marker(ctx, defs, line.tail, line.color, false) : "";
    return `<path d="${p.d}" fill="${fa}"${sa}${mk}/>`;
  });
  const flip: string[] = [];
  if (xf.fh) flip.push(`translate(${w.toFixed(2)} 0) scale(-1 1)`);
  if (xf.fv) flip.push(`translate(0 ${h.toFixed(2)}) scale(1 -1)`);
  const filt = sh && !noPaint ? `;filter:${shadowDrop(sh)}` : "";
  // the svg box covers the whole drawing (a callout's tail, arrow heads, half the stroke)
  const m = (line ? line.width / 2 : 0) + (line && (line.head || line.tail) ? line.width * 5 : 0);
  let bb = pathBBox(g.paths.map((p) => p.d).join(" "));
  if (bb && flip.length) {
    let [x0, y0, x1, y1] = bb;
    if (xf.fh) [x0, x1] = [w - x1, w - x0];
    if (xf.fv) [y0, y1] = [h - y1, h - y0];
    bb = [x0, y0, x1, y1];
  }
  const [bx, by, bx2, by2] = bb ?? [0, 0, w, h];
  const bx0 = Math.min(0, bx - m), by0 = Math.min(0, by - m);
  const bx1 = Math.max(Math.max(w, 1), bx2 + m), by1 = Math.max(Math.max(h, 1), by2 + m);
  const sw = bx1 - bx0, shh = by1 - by0;
  const rotCss = rot ? `;transform:rotate(${rot.toFixed(2)}deg);transform-origin:${(w / 2 - bx0).toFixed(1)}px ${(h / 2 - by0).toFixed(1)}px` : "";
  const move = bx0 || by0 ? `<g transform="translate(${(-bx0).toFixed(2)} ${(-by0).toFixed(2)})">` : "<g>";
  // the svg sits at 0,0 of a positioned holder: SVG rasterisers that copy the svg's computed
  // CSS into a picture would shift the drawing by its left/top; the rotation stays on the svg
  let svg = `<svg width="${sw.toFixed(2)}" height="${shh.toFixed(2)}" viewBox="0 0 ${sw.toFixed(2)} ${shh.toFixed(2)}" style="position:absolute;left:0;top:0;overflow:visible${filt}${rotCss}">` +
    (defs.length ? `<defs>${defs.join("")}</defs>` : "") + move + (flip.length ? `<g transform="${flip.join(" ")}">` : "<g>") + paths.join("") + "</g></g></svg>";
  if (bx0 || by0) svg = `<div style="position:absolute;left:${bx0.toFixed(2)}px;top:${by0.toFixed(2)}px;width:${sw.toFixed(2)}px;height:${shh.toFixed(2)}px">${svg}</div>`;
  // at least 1 px: a straight line's frame is 0 px high, and exporters skip such elements
  const st = [`position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${Math.max(w, 1).toFixed(1)}px;height:${Math.max(h, 1).toFixed(1)}px`];
  let inner = "";
  if (txBody && base) {
    let [l, t, r, b] = g.text;
    if (xf.fh) [l, r] = [w - r, w - l];
    if (xf.fv) [t, b] = [h - b, h - t];
    inner = textHtml(ctx, txBody, base, [l, t, r, b], chain, textShadow, [0, 0], rot ? [rot, w / 2, h / 2] : null);
  }
  return `<div style="${st.join(";")};">${svg}${inner}</div>`;
}

function missing(x: number, y: number, w: number, h: number): string {
  return `<div data-pptx-missing="${x.toFixed(1)},${y.toFixed(1)},${w.toFixed(1)},${h.toFixed(1)}" style="position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px"></div>`;
}

function cropCss(bf: El | null): string {
  const sr = find(bf, "a:srcRect");
  if (!sr) return "";
  const [l, t, r, b] = ["l", "t", "r", "b"].map((k) => int(attr(sr, k)) / 1000);
  if (!(l || t || r || b) || Math.min(l, t, r, b) < 0) return "";
  return `object-fit:fill;object-view-box:inset(${t.toFixed(3)}% ${r.toFixed(3)}% ${b.toFixed(3)}% ${l.toFixed(3)}%)`;
}

function renderPic(ctx: Ctx, el: El, space: Space, rels: Record<string, string>, onLayout: boolean): string {
  const spPr = kid(el, "spPr");
  const chain = phChain(ctx, phOf(el), onLayout);
  let xf = xfrmOf(spPr);
  for (const anc of chain) { if (xf) break; xf = xfrmOf(kid(anc, "spPr")); }
  if (!xf) return "";
  const [x, y] = space.point(xf.x, xf.y);
  const [w, h] = space.length(xf.w, xf.h);
  if (w <= 0 || h <= 0) return "";
  const bf = kid(el, "blipFill"), blip = find(bf, "a:blip");
  const uri = blipUri(ctx, blip, rels);
  if (!uri) { if (blip) { ctx.flags.add("bild"); return missing(x, y, w, h); } return ""; }
  const style = kid(el, "style");
  const line = lineOf(ctx, spPr, style, (Math.abs(space.sx) + Math.abs(space.sy)) / 2);
  const sh = shadowOf(ctx, spPr, style, ctx.scale);
  const geom = geomEl(spPr, chain);
  const prst = geom && ln(geom) === "prstGeom" ? attr(geom, "prst") ?? "rect" : "rect";
  const st = [`position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px`];
  const tf = [xf.rot ? `rotate(${xf.rot.toFixed(2)}deg)` : "", xf.fh ? "scaleX(-1)" : "", xf.fv ? "scaleY(-1)" : ""].filter(Boolean);
  if (tf.length) st.push(`transform:${tf.join(" ")}`);
  st.push(cropCss(bf) || "object-fit:fill");
  if (prst === "ellipse") st.push(Math.abs(w - h) < 0.5 ? `border-radius:${(w / 2).toFixed(1)}px` : "border-radius:50%");
  else if (prst === "roundRect") st.push(`border-radius:${((Math.max(0, Math.min(50000, adjust(geom, "adj", 16667))) / 100000) * Math.min(w, h)).toFixed(1)}px`);
  else if (prst !== "rect") { const g = shapePaths(geom, w, h); if (g && g.paths.length) st.push(`clip-path:path('${g.paths[0].d}')`); }
  // PowerPoint keeps the whole picture visible; only the outer half of its outline shows
  if (line) st.push(`outline:${(line.width / 2).toFixed(2)}px solid ${css(line.color)};outline-offset:0`);
  if (sh) st.push(["rect", "roundRect", "ellipse"].includes(prst) && !uri.startsWith("data:image/svg") ? shadowBox(sh) : `filter:${shadowDrop(sh)}`);
  const amt = find(blip, "a:alphaModFix");
  if (amt) st.push(`opacity:${(int(attr(amt, "amt"), 100000) / 100000).toFixed(3)}`);
  if (find(blip, "a:grayscl")) st.push("filter:grayscale(1)");
  return `<img src="${uri}" alt="" style="${st.filter(Boolean).join(";")};">`;
}

// ----------------------------------------------------------------------- tables
const MEDIUM2: Record<string, string> = {
  "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}": "accent1", "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}": "accent2",
  "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}": "accent3", "{00A15C55-8517-42AA-B614-E9B94910E393}": "accent4",
  "{7DF18680-E054-41AD-8BC1-D1AEF772440D}": "accent5", "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}": "accent6",
  "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}": "dk1",
};

/** "Medium Style 2", the table style every new PowerPoint table gets (often referenced by GUID only). */
export function medium2(accent: string): El {
  const lnx = (w: number) => `<a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
  const txt = '<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>';
  const solid = `<a:fill><a:solidFill><a:schemeClr val="${accent}"/></a:solidFill></a:fill>`;
  const tint = (v: number) => `<a:fill><a:solidFill><a:schemeClr val="${accent}"><a:tint val="${v}"/></a:schemeClr></a:solidFill></a:fill>`;
  return parseXml(`<a:tblStyle xmlns:a="${NS.a}">` +
    '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>' +
    "<a:tcStyle><a:tcBdr>" + ["left", "right", "top", "bottom", "insideH", "insideV"].map((s) => `<a:${s}>${lnx(12700)}</a:${s}>`).join("") +
    `</a:tcBdr>${tint(20000)}</a:tcStyle></a:wholeTbl>` +
    `<a:band1H><a:tcStyle><a:tcBdr/>${tint(40000)}</a:tcStyle></a:band1H><a:band1V><a:tcStyle><a:tcBdr/>${tint(40000)}</a:tcStyle></a:band1V>` +
    `<a:lastCol>${txt}<a:tcStyle><a:tcBdr/>${solid}</a:tcStyle></a:lastCol><a:firstCol>${txt}<a:tcStyle><a:tcBdr/>${solid}</a:tcStyle></a:firstCol>` +
    `<a:lastRow>${txt}<a:tcStyle><a:tcBdr><a:top>${lnx(38100)}</a:top></a:tcBdr>${solid}</a:tcStyle></a:lastRow>` +
    `<a:firstRow>${txt}<a:tcStyle><a:tcBdr><a:bottom>${lnx(38100)}</a:bottom></a:tcBdr>${solid}</a:tcStyle></a:firstRow></a:tblStyle>`);
}

function tableStyle(ctx: Ctx, sid: string | null): El | null {
  const id = sid || ctx.pkg.defaultTableStyle;
  if (!id) return null;
  return ctx.pkg.tableStyles[id] ?? (MEDIUM2[id] ? medium2(MEDIUM2[id]) : null);
}

function renderTable(ctx: Ctx, el: El, space: Space): string {
  const xfrm = kid(el, "xfrm"), tbl = find(el, ".//a:tbl");
  const off = find(xfrm, "a:off");
  if (!xfrm || !tbl || !off) return "";
  const [x, y] = space.point(parseFloat(attr(off, "x") ?? "0"), parseFloat(attr(off, "y") ?? "0"));
  const s = ctx.scale;
  const tblPr = find(tbl, "a:tblPr");
  const flag = (k: string) => ["1", "true"].includes(attr(tblPr, k) ?? "");
  const style = tableStyle(ctx, textOf(find(tblPr, "a:tableStyleId")) || null);
  const cols = findAll(tbl, "a:tblGrid/a:gridCol").map((g) => int(attr(g, "w")) * space.sx);
  const rows = findAll(tbl, "a:tr");
  const nr = rows.length, nc = cols.length;
  const partsFor = (r: number, c: number): El[] => {
    if (!style) return [];
    const names = ["wholeTbl"];
    if (flag("bandRow") && !(flag("firstRow") && r === 0) && !(flag("lastRow") && r === nr - 1))
      names.push((r - (flag("firstRow") ? 1 : 0)) % 2 === 0 ? "band1H" : "band2H");
    if (flag("bandCol")) names.push((c - (flag("firstCol") ? 1 : 0)) % 2 === 0 ? "band1V" : "band2V");
    if (flag("firstCol") && c === 0) names.push("firstCol");
    if (flag("lastCol") && c === nc - 1) names.push("lastCol");
    if (flag("lastRow") && r === nr - 1) names.push("lastRow");
    if (flag("firstRow") && r === 0) names.push("firstRow");
    return names.map((n) => find(style, `a:${n}`)).filter((e): e is El => !!e);
  };
  const border = (l: El | null): string | null => {
    if (!l) return null;
    if (find(l, "a:noFill")) return "none";
    const c = fillEl(l) ? ctx.colorIn(fillEl(l)) : null;
    return c ? `${Math.max(0.75, int(attr(l, "w"), 12700) * s).toFixed(2)}px solid ${css(c)}` : null;
  };
  const htmlRows = rows.map((tr, r) => {
    const rh = int(attr(tr, "h")) * space.sy;
    const cells: string[] = [];
    let c = 0;
    for (const tc of findAll(tr, "a:tc")) {
      if (["1", "true"].includes(attr(tc, "hMerge") ?? "") || ["1", "true"].includes(attr(tc, "vMerge") ?? "")) { c++; continue; }
      let fill: Fill | null = null;
      const bd: Record<string, string | null> = { left: null, right: null, top: null, bottom: null };
      let txColor: RGBA | null = null, txBold: boolean | null = null, txFont: string | null = null;
      for (const part of partsFor(r, c)) {
        const tcs = find(part, "a:tcStyle");
        const f = find(tcs, "a:fill");
        if (f && fillEl(f)) fill = fillOf(ctx, fillEl(f));
        const bdr = find(tcs, "a:tcBdr");
        const outer = ["firstRow", "lastRow", "firstCol", "lastCol"].includes(ln(part));
        if (bdr) for (const [side, inner, edge] of [["left", "insideV", c === 0], ["right", "insideV", c === nc - 1],
                                                   ["top", "insideH", r === 0], ["bottom", "insideH", r === nr - 1]] as [string, string, boolean][]) {
          let src = edge || outer ? find(bdr, `a:${side}/a:ln`) : null;
          if (!src && !edge && !outer) src = find(bdr, `a:${inner}/a:ln`);
          const v = border(src);
          if (v !== null) bd[side] = v;
        }
        const tts = find(part, "a:tcTxStyle");
        if (tts) {
          const col = ctx.colorIn(tts);
          if (col) txColor = col;
          if (attr(tts, "b")) txBold = attr(tts, "b") === "on";
          const fr = find(tts, "a:fontRef");
          if (fr) txFont = attr(fr, "idx");
        }
      }
      const tcPr = find(tc, "a:tcPr");
      const ownFill = fillOf(ctx, fillEl(tcPr));
      if (ownFill) fill = ownFill;
      for (const [side, tag] of [["left", "lnL"], ["right", "lnR"], ["top", "lnT"], ["bottom", "lnB"]]) {
        const v = border(find(tcPr, `a:${tag}`));
        if (v !== null) bd[side] = v;
      }
      const mar = (k: string, d: number) => int(attr(tcPr, k), d) * s;
      const half = (side: string) => { const m = (bd[side] ?? "").match(/^([\d.]+)px/); return m ? parseFloat(m[1]) / 2 : 0; };
      // collapsed borders sit centred on the grid line; PowerPoint's margin starts at the line
      const cst = [`padding:${Math.max(0, mar("marT", 45720) - half("top")).toFixed(1)}px ${Math.max(0, mar("marR", 91440) - half("right")).toFixed(1)}px ${Math.max(0, mar("marB", 45720) - half("bottom")).toFixed(1)}px ${Math.max(0, mar("marL", 91440) - half("left")).toFixed(1)}px`,
        `vertical-align:${({ ctr: "middle", b: "bottom" } as Record<string, string>)[attr(tcPr, "anchor") ?? "t"] ?? "top"}`, "box-sizing:border-box"];
      if (fill && fill.kind !== "none") cst.push(fill.css());
      for (const [side, v] of Object.entries(bd)) if (v && v !== "none") cst.push(`border-${side}:${v}`);
      let span = "";
      if (int(attr(tc, "gridSpan"), 1) > 1) span += ` colspan="${int(attr(tc, "gridSpan"))}"`;
      if (int(attr(tc, "rowSpan"), 1) > 1) span += ` rowspan="${int(attr(tc, "rowSpan"))}"`;
      let inner = "";
      const tb = find(tc, "a:txBody");
      if (tb) {
        const base = new TextBase([lstLevels(ctx.pkg.defaultLst), lstLevels(ctx.txOther)], txColor, txFont);
        base.fontrefAt = 2;
        const [h, cellCss] = cellText(ctx, tb, base, txBold);
        inner = h;
        cst.push(...cellCss); // exporters read alignment and type from the <td>
      }
      cells.push(`<td${span} style="${cst.join(";")}">${inner}</td>`);
      c += int(attr(tc, "gridSpan"), 1);
    }
    return `<tr style="height:${rh.toFixed(1)}px">${cells.join("")}</tr>`;
  });
  const colgroup = "<colgroup>" + cols.map((w) => `<col style="width:${w.toFixed(1)}px">`).join("") + "</colgroup>";
  const total = cols.reduce((a, b) => a + b, 0);
  return `<table style="position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${total.toFixed(1)}px;border-collapse:collapse;table-layout:fixed;border-spacing:0">${colgroup}${htmlRows.join("")}</table>`;
}

function cellText(ctx: Ctx, txBody: El, base: TextBase, bold: boolean | null): [string, string[]] {
  const out: string[] = [];
  let tdCss: string[] = [];
  findAll(txBody, "a:p").forEach((p, i) => {
    const pPr = find(p, "a:pPr");
    const props = pprMerge(base.level(pPr ? int(attr(pPr, "lvl")) + 1 : 1), pPr);
    if (bold !== null) props.rpr = { ...props.rpr, b: bold ? "1" : "0" };
    const spans: string[] = [], sizes: number[] = [];
    let first: Run | null = null;
    for (const r of kids(p)) {
      const t = ln(r);
      if (t === "br") spans.push("<br>");
      else if (t === "r" || t === "fld") {
        const rp = runProps(ctx, rprMerge(props.rpr ?? {}, find(r, "a:rPr")), 1, base.fontRef);
        sizes.push(rp.size);
        first = first ?? rp;
        const txt = textOf(find(r, "a:t"));
        if (txt) spans.push(`<span style="${runCss(rp, ctx)}">${esc(txt)}</span>`);
      }
    }
    const endP = runProps(ctx, rprMerge(props.rpr ?? {}, find(p, "a:endParaRPr")), 1, base.fontRef);
    const size = sizes.length ? Math.max(...sizes) : endP.size;
    const algn = ({ ctr: "center", r: "right", just: "justify" } as Record<string, string>)[props.algn ?? "l"] ?? "left";
    const st = [`line-height:${(1.2 * size).toFixed(2)}px`, `font-size:${size.toFixed(1)}px`, "margin:0", "white-space:pre-wrap"];
    if (algn !== "left") st.push(`text-align:${algn}`);
    if (i === 0) {
      const fp: Run = first ?? endP;
      tdCss = [`text-align:${algn}`, `font-size:${fp.size.toFixed(1)}px`, faceCss(fp.face), `color:${css(fp.color)}`, ...(fp.bold ? ["font-weight:700"] : [])];
      if (first) { const shf = ctx.firstLineShift(first.face, size, 1); if (Math.abs(shf) >= 0.05) st.push(`position:relative;top:${shf.toFixed(1)}px`); }
    }
    out.push(`<div style="${st.join(";")}">${spans.join("") || "&#8203;"}</div>`);
  });
  return [out.join(""), tdCss];
}

// ----------------------------------------------------------------------- trees
const MC_KNOWN = new Set(["p14", "a14", "p15", "a15", "a16", "p16", "v", "o", "p159", "p1710", "a1611", "ahyp", "asvg", "p", "a", "r", "mc", "wps", "wpg"]);

function renderFrame(ctx: Ctx, el: El, space: Space, rels: Record<string, string>): string {
  const gd = find(el, "a:graphic/a:graphicData");
  const uri = attr(gd, "uri") ?? "";
  const xf = kid(el, "xfrm");
  if (uri.endsWith("/table")) return renderTable(ctx, el, space);
  if (uri.endsWith("/diagram")) {
    // SmartArt: PowerPoint stores a pre-rendered drawing next to the data model
    for (const [, typ, path] of ctx.pkg.rels(ctx.curPart)) {
      if (!typ.endsWith("/diagramDrawing") || !ctx.pkg.has(path)) continue;
      const drawing = ctx.pkg.xml(path);
      const tree = find(drawing, ".//dsp:spTree") ?? findAll(drawing, ".//*").find((e) => ln(e) === "spTree") ?? null;
      const off = find(xf, "a:off");
      if (!tree || !off) return "";
      const [ox, oy] = space.point(parseFloat(attr(off, "x") ?? "0"), parseFloat(attr(off, "y") ?? "0"));
      const prev = ctx.curPart;
      ctx.curPart = path;
      try { return renderTree(ctx, tree, new Space(space.sx, space.sy, ox, oy), ctx.pkg.relMap(path)).join(""); }
      finally { ctx.curPart = prev; }
    }
    return "";
  }
  if (uri.endsWith("/ole")) {
    const pic = findAll(gd, ".//*").find((e) => ln(e) === "pic");
    if (pic) return renderPic(ctx, pic, space, rels, false);
  }
  if (uri.endsWith("/chart")) {
    const ref = kids(gd).find((c) => ln(c) === "chart") ?? null;
    const path = rels[rId(ref, "id") ?? ""];
    const off = find(xf, "a:off"), ext = find(xf, "a:ext");
    if (path && ctx.pkg.has(path) && off && ext) {
      const [x, y] = space.point(parseFloat(attr(off, "x") ?? "0"), parseFloat(attr(off, "y") ?? "0"));
      const [w, h] = space.length(parseFloat(attr(ext, "cx") ?? "0"), parseFloat(attr(ext, "cy") ?? "0"));
      const color = (spPr: El | null): RGBA | null => {
        const f = fillOf(ctx, fillEl(spPr));
        if (!f || f.kind === "none") return null;
        return f.kind === "solid" ? f.color : f.stops[0]?.[1] ?? null;
      };
      const face = ctx.font(null, "minor");
      ctx.pkg.faces.add(face);
      const svg = renderChart(ctx.pkg.xml(path), w, h, color, [1, 2, 3, 4, 5, 6].map((i) => ctx.scheme(`accent${i}`)), face, 2 * ctx.scale * EMU_PER_PX);
      if (svg) return `<div data-pptx-chart="1" style="position:absolute;left:${x.toFixed(1)}px;top:${y.toFixed(1)}px;width:${w.toFixed(1)}px;height:${h.toFixed(1)}px;">${svg}</div>`;
    }
  }
  ctx.flags.add(uri.endsWith("/chart") ? "diagramm" : "objekt");
  const off = find(xf, "a:off"), ext = find(xf, "a:ext");
  if (!off || !ext) return "";
  const [x, y] = space.point(parseFloat(attr(off, "x") ?? "0"), parseFloat(attr(off, "y") ?? "0"));
  const [w, h] = space.length(parseFloat(attr(ext, "cx") ?? "0"), parseFloat(attr(ext, "cy") ?? "0"));
  return missing(x, y, w, h);
}

function renderGroup(ctx: Ctx, el: El, space: Space, rels: Record<string, string>, groupFill: Fill | null, onLayout: boolean): string[] {
  const grpPr = kid(el, "grpSpPr");
  const xf = find(grpPr, "a:xfrm");
  if (!xf) return renderTree(ctx, el, space, rels, groupFill, onLayout);
  const off = find(xf, "a:off"), ext = find(xf, "a:ext");
  if (!off || !ext) return [];
  const choff = find(xf, "a:chOff") ?? off, chext = find(xf, "a:chExt") ?? ext;
  const v = (e: El, k: string) => parseFloat(attr(e, k) ?? "0");
  const gfill = fillOf(ctx, fillEl(grpPr), null, rels) ?? groupFill;
  const rot = int(attr(xf, "rot")) / 60000;
  const fh = ["1", "true"].includes(attr(xf, "flipH") ?? ""), fv = ["1", "true"].includes(attr(xf, "flipV") ?? "");
  if (!rot && !fh && !fv)
    return renderTree(ctx, el, space.child([v(off, "x"), v(off, "y")], [v(ext, "cx"), v(ext, "cy")], [v(choff, "x"), v(choff, "y")], [v(chext, "cx"), v(chext, "cy")]), rels, gfill, onLayout);
  // rotated/flipped group: one wrapper element carries the transform
  const [gx, gy] = space.point(v(off, "x"), v(off, "y"));
  const [gw, gh] = space.length(v(ext, "cx"), v(ext, "cy"));
  const local = new Space(gw / (v(chext, "cx") || 1), gh / (v(chext, "cy") || 1), 0, 0, v(choff, "x"), v(choff, "y"));
  const tf = [rot ? `rotate(${rot.toFixed(2)}deg)` : "", fh ? "scaleX(-1)" : "", fv ? "scaleY(-1)" : ""].filter(Boolean).join(" ");
  return [`<div style="position:absolute;left:${gx.toFixed(1)}px;top:${gy.toFixed(1)}px;width:${gw.toFixed(1)}px;height:${gh.toFixed(1)}px;transform:${tf}">${renderTree(ctx, el, local, rels, gfill, onLayout).join("")}</div>`];
}

function renderTree(ctx: Ctx, tree: El, space: Space, rels: Record<string, string>, groupFill: Fill | null = null,
                    onLayout = false, skipPlaceholders = false): string[] {
  const out: string[] = [];
  for (const el of kids(tree)) {
    const tag = ln(el);
    if (tag === "AlternateContent") {
      const choice = find(el, "mc:Choice");
      const req = new Set((attr(choice, "Requires") ?? "").split(/\s+/).filter(Boolean));
      const pick = choice && [...req].every((r) => MC_KNOWN.has(r)) ? choice : find(el, "mc:Fallback");
      if (pick) out.push(...renderTree(ctx, pick, space, rels, groupFill, onLayout, skipPlaceholders));
      continue;
    }
    if (skipPlaceholders && ["sp", "pic", "graphicFrame"].includes(tag) && phOf(el)) continue;
    try {
      if (tag === "sp") out.push(renderSp(ctx, el, space, rels, groupFill, onLayout, false));
      else if (tag === "cxnSp") out.push(renderSp(ctx, el, space, rels, groupFill, onLayout, true));
      else if (tag === "pic") out.push(renderPic(ctx, el, space, rels, onLayout));
      else if (tag === "graphicFrame") out.push(renderFrame(ctx, el, space, rels));
      else if (tag === "grpSp") out.push(...renderGroup(ctx, el, space, rels, groupFill, onLayout));
    } catch { /* one odd shape must not cost the slide */ }
  }
  return out.filter(Boolean);
}

function bgFill(ctx: Ctx, root: El | null, rels: Record<string, string>): Fill | null {
  const bg = find(root, "p:cSld/p:bg");
  if (!bg) return null;
  const bgPr = find(bg, "p:bgPr");
  if (bgPr) return fillOf(ctx, fillEl(bgPr), null, rels);
  const ref = find(bg, "p:bgRef");
  if (ref) {
    const idx = int(attr(ref, "idx")), c = ctx.colorIn(ref);
    const [lst, i] = idx >= 1001 ? [ctx.theme.bgfills, idx - 1001] : [ctx.theme.fills, idx - 1];
    if (i >= 0 && i < lst.length) return fillOf(ctx, lst[i], c, rels);
    return c ? new Fill("solid", c) : null;
  }
  return null;
}

export interface SlideResult { html: string; flags: string[]; number: number }

export function renderSlide(pkg: Pkg, slidePath: string, root: Space, no: number, opts: ConvertOptions): SlideResult {
  const ctx = new Ctx(pkg, slidePath, root.sx, opts);
  ctx.slideNo = no;
  const sRels = pkg.relMap(slidePath), lRels = pkg.relMap(ctx.layoutPath), mRels = pkg.relMap(ctx.masterPath);
  let bg: Fill | null = null;
  for (const [r, rl] of [[ctx.slide, sRels], [ctx.layout, lRels], [ctx.master, mRels]] as [El | null, Record<string, string>][]) {
    bg = bgFill(ctx, r, rl);
    if (bg) break;
  }
  let bgCss = (bg && bg.kind !== "none" ? bg.css() : "") || "background:#FFFFFF";
  if (bg && bg.kind === "img") bgCss = bgCss.replace("background-size:100% 100%", "background-size:cover");
  const parts: string[] = [];
  const showMaster = !["0", "false"].includes(attr(ctx.slide, "showMasterSp") ?? "");
  const layoutShow = !ctx.layout || !["0", "false"].includes(attr(ctx.layout, "showMasterSp") ?? "");
  if (showMaster && layoutShow && ctx.master) {
    const tree = find(ctx.master, "p:cSld/p:spTree");
    if (tree) { ctx.curPart = ctx.masterPath!; parts.push(...renderTree(ctx, tree, root, mRels, null, true, true)); }
  }
  if (showMaster && ctx.layout) {
    const tree = find(ctx.layout, "p:cSld/p:spTree");
    if (tree) { ctx.curPart = ctx.layoutPath!; parts.push(...renderTree(ctx, tree, root, lRels, null, true, true)); }
  }
  const tree = find(ctx.slide, "p:cSld/p:spTree");
  if (tree) { ctx.curPart = slidePath; parts.push(...renderTree(ctx, tree, root, sRels)); }
  const flags = [...ctx.flags].sort();
  const fl = flags.length ? ` data-pptx-flags="${flags.join(",")}"` : "";
  const extra = opts.sectionAttributes ? ` ${opts.sectionAttributes}` : "";
  return { number: no, flags,
           html: `<section class="slide" data-pptx-import="1"${fl}${extra} style="${bgCss};">${parts.join("")}</section>` };
}

export { fitSpace, Space };
