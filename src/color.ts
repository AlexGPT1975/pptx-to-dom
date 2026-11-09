// Colours: DrawingML colour elements with their modifiers, CSS output, gradients blended in
// linear light (as PowerPoint does) and pattern fills as SVG tiles.
import { El, attr, find, int, kids, ln } from "./xml.js";

export type RGBA = [number, number, number, number]; // r/g/b 0..255, a 0..1

export const PRESET_COLORS: Record<string, string> = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
  gray: "808080", grey: "808080", darkGray: "A9A9A9", lightGray: "D3D3D3", orange: "FFA500", purple: "800080",
  navy: "000080", silver: "C0C0C0", maroon: "800000", teal: "008080", olive: "808000", lime: "00FF00",
  aqua: "00FFFF", fuchsia: "FF00FF",
};
export const COLOR_TAGS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"]);

export function hex(v: string | null | undefined): RGBA {
  let s = (v ?? "000000").trim().replace(/^#/, "");
  if (s.length !== 6) s = "000000";
  return [parseInt(s.slice(0, 2), 16), parseInt(s.slice(2, 4), 16), parseInt(s.slice(4, 6), 16), 1];
}

export const toLin = (c: number): number => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
export const fromLin = (c: number): number => {
  c = Math.max(0, Math.min(1, c));
  return 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055);
};

function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  const l = (mx + mn) / 2;
  if (mx === mn) return [0, l, 0];
  const d = mx - mn;
  const s = l <= 0.5 ? d / (mx + mn) : d / (2 - mx - mn);
  let h: number;
  if (mx === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (mx === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, l, s];
}
function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) return [l, l, l];
  const m2 = l <= 0.5 ? l * (1 + s) : l + s - l * s, m1 = 2 * l - m2;
  const v = (hue: number) => {
    hue = ((hue % 1) + 1) % 1;
    if (hue < 1 / 6) return m1 + (m2 - m1) * hue * 6;
    if (hue < 0.5) return m2;
    if (hue < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - hue) * 6;
    return m1;
  };
  return [v(h + 1 / 3), v(h), v(h - 1 / 3)];
}

export function applyMods(c: RGBA, el: El): RGBA {
  let [r, g, b, a] = c;
  for (const m of kids(el)) {
    const t = ln(m);
    const v = int(attr(m, "val"), 100000) / 100000;
    if (t === "alpha") a = v;
    else if (t === "alphaMod") a *= v;
    else if (t === "alphaOff") a += v;
    else if (t === "tint" || t === "shade") {
      const lin = [r, g, b].map(toLin).map((x) => (t === "tint" ? x * v + (1 - v) : x * v));
      [r, g, b] = lin.map(fromLin);
    } else if (["lumMod", "lumOff", "lum", "satMod", "satOff", "sat", "hueMod", "hueOff", "hue"].includes(t)) {
      let [h, l, s] = rgbToHls(r / 255, g / 255, b / 255);
      if (t === "lumMod") l *= v;
      else if (t === "lumOff") l += v;
      else if (t === "lum") l = v;
      else if (t === "satMod") s *= v;
      else if (t === "satOff") s += v;
      else if (t === "sat") s = v;
      else if (t === "hueMod") h = (h * v) % 1;
      else if (t === "hueOff") h = (h + int(attr(m, "val")) / 21600000) % 1;
      else if (t === "hue") h = int(attr(m, "val")) / 21600000;
      l = Math.max(0, Math.min(1, l)); s = Math.max(0, Math.min(1, s));
      [r, g, b] = hlsToRgb(h, l, s).map((x) => x * 255);
    } else if (t === "comp") {
      const [h, l, s] = rgbToHls(r / 255, g / 255, b / 255);
      [r, g, b] = hlsToRgb((h + 0.5) % 1, l, s).map((x) => x * 255);
    } else if (t === "inv") {
      [r, g, b] = [255 - r, 255 - g, 255 - b];
    } else if (t === "gray") {
      const y = 0.3 * r + 0.59 * g + 0.11 * b;
      r = g = b = y;
    }
  }
  const cl = (x: number) => Math.max(0, Math.min(255, x));
  return [cl(r), cl(g), cl(b), Math.max(0, Math.min(1, a))];
}

const h2 = (x: number) => Math.round(x).toString(16).toUpperCase().padStart(2, "0");

export function css(c: RGBA | null | undefined): string | null {
  if (!c) return null;
  const [r, g, b, a] = c;
  if (a >= 0.999) return `#${h2(r)}${h2(g)}${h2(b)}`;
  return `rgba(${Math.round(r)},${Math.round(g)},${Math.round(b)},${a.toFixed(3)})`;
}
export const opaque = (c: RGBA): RGBA => [c[0], c[1], c[2], 1];

/** One theme part: colours, fonts, format scheme (fills/lines/effects/bg fills). */
export class Theme {
  colors: Record<string, string> = {};
  major = "";
  minor = "";
  fills: El[] = [];
  lines: El[] = [];
  effects: El[] = [];
  bgfills: El[] = [];

  constructor(root: El | null) {
    if (!root) return;
    const scheme = find(root, ".//a:clrScheme");
    for (const c of kids(scheme)) {
      const srgb = find(c, "a:srgbClr"), sys = find(c, "a:sysClr");
      if (srgb) this.colors[ln(c)] = attr(srgb, "val") ?? "000000";
      else if (sys) this.colors[ln(c)] = attr(sys, "lastClr") ?? (attr(sys, "val") === "window" ? "FFFFFF" : "000000");
    }
    this.major = attr(find(root, ".//a:fontScheme/a:majorFont/a:latin"), "typeface") ?? "";
    this.minor = attr(find(root, ".//a:fontScheme/a:minorFont/a:latin"), "typeface") ?? "";
    const fs = find(root, ".//a:fmtScheme");
    this.fills = kids(find(fs, "a:fillStyleLst"));
    this.lines = kids(find(fs, "a:lnStyleLst"));
    this.effects = kids(find(fs, "a:effectStyleLst"));
    this.bgfills = kids(find(fs, "a:bgFillStyleLst"));
  }
}

/** PowerPoint blends gradient stops in linear light; a browser blends in sRGB, which makes
 * red-to-pink or dark-to-light gradients visibly darker in the middle. Intermediate stops
 * computed in linear light make any renderer draw PowerPoint's curve. */
export function linearLight(stops: [number, RGBA][], steps = 6): [number, RGBA][] {
  if (stops.length < 2) return stops;
  const out: [number, RGBA][] = [stops[0]];
  for (let i = 0; i + 1 < stops.length; i++) {
    const [p0, c0] = stops[i], [p1, c1] = stops[i + 1];
    if (c0[0] !== c1[0] || c0[1] !== c1[1] || c0[2] !== c1[2]) {
      for (let k = 1; k < steps; k++) {
        const u = k / steps;
        const rgb = [0, 1, 2].map((j) => fromLin(toLin(c0[j]) * (1 - u) + toLin(c1[j]) * u));
        out.push([p0 + (p1 - p0) * u, [rgb[0], rgb[1], rgb[2], c0[3] + (c1[3] - c0[3]) * u]]);
      }
    }
    out.push([p1, c1]);
  }
  return out;
}

export function shadeColor(c: RGBA | null, mode: string | null): RGBA | null {
  if (!c || !mode) return c;
  const k = ({ darken: 0.6, darkenLess: 0.8, lighten: 0.6, lightenLess: 0.8 } as Record<string, number>)[mode];
  if (k === undefined) return c;
  const lin = [c[0], c[1], c[2]].map(toLin).map((x) => (mode.startsWith("darken") ? x * k : x * k + (1 - k)));
  return [fromLin(lin[0]), fromLin(lin[1]), fromLin(lin[2]), c[3]];
}

export function base64(data: Uint8Array | string): string {
  const bytes = typeof data === "string" ? new TextEncoder().encode(data) : data;
  const B = (globalThis as { Buffer?: { from(b: Uint8Array): { toString(enc: string): string } } }).Buffer;
  if (B) return B.from(bytes).toString("base64");
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

// PowerPoint's pattern fills are 8x8-pixel tiles (at 96 dpi; 1.5 x on a 144-dpi 1920 stage).
const BAYER = [[0, 32, 8, 40, 2, 34, 10, 42], [48, 16, 56, 24, 50, 18, 58, 26], [12, 44, 4, 36, 14, 46, 6, 38],
  [60, 28, 52, 20, 62, 30, 54, 22], [3, 35, 11, 43, 1, 33, 9, 41], [51, 19, 59, 27, 49, 17, 57, 25],
  [15, 47, 7, 39, 13, 45, 5, 37], [63, 31, 55, 23, 61, 29, 53, 21]];
const md = (a: number, n: number) => ((a % n) + n) % n;
export const PATTERNS: Record<string, (x: number, y: number) => boolean> = {
  horz: (x, y) => y % 8 === 0, vert: (x) => x % 8 === 0, ltHorz: (x, y) => y % 4 === 0, ltVert: (x) => x % 4 === 0,
  dkHorz: (x, y) => y % 4 < 2, dkVert: (x) => x % 4 < 2, narHorz: (x, y) => y % 2 === 0, narVert: (x) => x % 2 === 0,
  wdHorz: (x, y) => y % 8 < 3, wdVert: (x) => x % 8 < 3,
  dnDiag: (x, y) => md(x - y, 8) === 0, upDiag: (x, y) => (x + y) % 8 === 0,
  ltDnDiag: (x, y) => md(x - y, 4) === 0, ltUpDiag: (x, y) => (x + y) % 4 === 0,
  dkDnDiag: (x, y) => md(x - y, 4) < 2, dkUpDiag: (x, y) => (x + y) % 4 < 2,
  wdDnDiag: (x, y) => md(x - y, 8) < 3, wdUpDiag: (x, y) => (x + y) % 8 < 3,
  cross: (x, y) => x % 8 === 0 || y % 8 === 0, smGrid: (x, y) => x % 4 === 0 || y % 4 === 0,
  lgGrid: (x, y) => x % 8 === 0 || y % 8 === 0, diagCross: (x, y) => md(x - y, 8) === 0 || (x + y) % 8 === 0,
  smCheck: (x, y) => (Math.floor(x / 2) + Math.floor(y / 2)) % 2 === 0,
  lgCheck: (x, y) => (Math.floor(x / 4) + Math.floor(y / 4)) % 2 === 0,
  dashHorz: (x, y) => y % 4 === 0 && x % 8 < 4, dashVert: (x, y) => x % 4 === 0 && y % 8 < 4,
  dotGrid: (x, y) => x % 4 === 0 && y % 4 === 0,
};
for (const p of [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90]) {
  PATTERNS[`pct${p}`] = (x, y) => BAYER[y % 8][x % 8] < (p * 64) / 100; // ordered dither
}

export function patternSvg(prst: string, fg: RGBA, bg: RGBA, px: number): string | null {
  const on = PATTERNS[prst];
  if (!on) return null;
  let rects = "";
  for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++)
    if (on(x, y)) rects += `<rect x="${(x * px).toFixed(2)}" y="${(y * px).toFixed(2)}" width="${px.toFixed(2)}" height="${px.toFixed(2)}"/>`;
  const size = (8 * px).toFixed(2);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" shape-rendering="crispEdges">` +
    `<rect width="100%" height="100%" fill="${css(opaque(bg))}" fill-opacity="${bg[3].toFixed(3)}"/>` +
    `<g fill="${css(opaque(fg))}" fill-opacity="${fg[3].toFixed(3)}">${rects}</g></svg>`;
  return "data:image/svg+xml;base64," + base64(svg);
}

