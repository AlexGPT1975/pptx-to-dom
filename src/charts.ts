// PPTX charts -> SVG, from the chart part's own cached data (c:numCache / c:strCache).
// PowerPoint stores no picture of a chart, only its definition and the last values. Covered:
// bar/column (clustered, stacked, 100 %), line, area, pie, doughnut - with title, legend,
// category and value axes, gridlines and data labels. Every label stays real <text>.
import { El, attr, find, findAll, int, kids, ln, textOf } from "./xml.js";
import { RGBA } from "./color.js";

export type ColorFn = (spPr: El | null) => RGBA | null;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const f1 = (v: number) => v.toFixed(1);

function cache(ref: El | null): string[] {
  if (!ref) return [];
  const c = findAll(ref, ".//c:strCache")[0] ?? findAll(ref, ".//c:numCache")[0];
  if (!c) return findAll(ref, ".//c:pt").map((p) => textOf(find(p, "c:v")));
  const count = int(attr(find(c, "c:ptCount"), "val"));
  const pts: Record<number, string> = {};
  let max = -1;
  for (const pt of findAll(c, "c:pt")) {
    const i = int(attr(pt, "idx"));
    pts[i] = textOf(find(pt, "c:v"));
    max = Math.max(max, i);
  }
  const n = Math.max(count, max + 1);
  return Array.from({ length: n }, (_, i) => pts[i] ?? "");
}

const toNum = (v: string): number | null => { const f = parseFloat(v); return isNaN(f) ? null : f; };

function richText(rich: El | null): string {
  if (!rich) return "";
  return findAll(rich, ".//a:p").map((p) => findAll(p, ".//a:t").map(textOf).join("")).join(" ").trim();
}

/** A value with its number format (decimals, %), German separators. */
export function fmtValue(v: number, codeIn: string): string {
  const code = (codeIn || "General").split(";")[0];
  const pct = code.includes("%");
  if (pct) v *= 100;
  let dec: number;
  if (code !== "General" && code.includes(".")) {
    dec = 0;
    for (const ch of code.split(".", 2)[1]) { if (ch === "0" || ch === "#") dec++; else break; }
  } else if (code !== "General") dec = 0;
  else dec = Math.abs(v - Math.round(v)) < 1e-9 ? 0 : Math.abs(v) >= 10 ? 1 : 2;
  const s = v.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec })
    .replace(/,/g, "X").replace(/\./g, ",").replace(/X/g, ".");
  return s + (pct ? " %" : "");
}

function paint(c: RGBA | null): string {
  if (!c) return "none";
  const [r, g, b, a] = c;
  const h = (x: number) => Math.round(x).toString(16).toUpperCase().padStart(2, "0");
  return a < 0.999 ? `rgba(${Math.round(r)},${Math.round(g)},${Math.round(b)},${a.toFixed(3)})` : `#${h(r)}${h(g)}${h(b)}`;
}

interface Series { name: string; cats: string[]; vals: (number | null)[]; color: RGBA; pts: Record<number, RGBA>;
                   code: string; labels: boolean; marker: string | null }

function labelsOn(ser: El, pl: El): boolean {
  for (const src of [find(ser, "c:dLbls"), find(pl, "c:dLbls")]) {
    const sv = find(src, "c:showVal");
    if (sv) return ["1", "true"].includes(attr(sv, "val") ?? "1");
  }
  return false;
}

function markerOf(ser: El, pl: El): string | null {
  if (!["lineChart", "line3DChart"].includes(ln(pl))) return null;
  const sym = find(ser, "c:marker/c:symbol");
  if (sym) { const v = attr(sym, "val") ?? "none"; return v === "none" ? null : v === "square" ? "square" : "circle"; }
  const mk = find(pl, "c:marker");
  return mk && ["1", "true"].includes(attr(mk, "val") ?? "1") ? "circle" : null;
}

/** PowerPoint's automatic value axis: 5 % headroom beyond the data, a major unit of 1/2/5 x 10^n
 * giving about seven intervals, both ends rounded out to the unit. */
export function niceAxis(lo: number, hi: number): [number, number, number] {
  if (hi === lo) hi = lo + 1;
  const span = hi - lo;
  const hi2 = hi > 0 ? hi + 0.05 * span : hi;
  const lo2 = lo < 0 ? lo - 0.05 * span : lo;
  const raw = (hi2 - lo2) / 7;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 5, 10].map((m) => m * mag).find((s) => s >= raw - 1e-12) ?? 10 * mag;
  return [Math.floor(lo2 / step) * step, Math.ceil(hi2 / step - 1e-9) * step, step];
}

export function renderChart(chart: El, w: number, h: number, color: ColorFn, accents: RGBA[], face: string, scale: number): string {
  const plot = find(chart, ".//c:plotArea");
  if (!plot) return "";
  const kinds = ["barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "pieChart", "pie3DChart", "doughnutChart"];
  const pl = kids(plot).find((e) => kinds.includes(ln(e)));
  if (!pl) return "";
  const kind = ln(pl).replace("3D", "");
  let fs = 12 * scale * 0.9;
  const txpr = find(chart, "c:txPr") ? findAll(find(chart, "c:txPr"), ".//a:defRPr")[0] : null;
  if (txpr && attr(txpr, "sz")) fs = (int(attr(txpr, "sz")) / 100) * scale;
  const font = `font-family:'${face}',Arial,sans-serif`;
  const out: string[] = [];
  const titleEl = find(chart, "c:chart/c:title");
  const autoDel = find(chart, "c:chart/c:autoTitleDeleted");
  const sers = findAll(pl, "c:ser");
  let title = titleEl ? richText(find(titleEl, "c:tx/c:rich")) : "";
  if (titleEl && !title && sers.length === 1 && attr(autoDel, "val") !== "1")
    title = richText(find(sers[0], "c:tx")) || cache(find(sers[0], "c:tx/c:strRef")).join(" ");
  let top = 8;
  if (title) {
    const tr = titleEl ? findAll(titleEl, ".//a:defRPr")[0] : null;
    const tsz = tr && attr(tr, "sz") ? (int(attr(tr, "sz")) / 100) * scale : fs * 1.2;
    out.push(`<text x="${f1(w / 2)}" y="${f1(top + tsz)}" text-anchor="middle" font-size="${f1(tsz)}" fill="#595959" style="${font}">${esc(title)}</text>`);
    top += tsz * 1.6;
  }
  const series: Series[] = sers.map((s, i) => {
    const name = cache(find(s, "c:tx/c:strRef")).join(" ") || richText(find(s, "c:tx")) || `Reihe ${i + 1}`;
    const catRef = find(s, "c:cat/c:strRef") ?? find(s, "c:cat/c:numRef") ?? find(s, "c:cat/c:multiLvlStrRef");
    const idx = find(s, "c:idx") ? int(attr(find(s, "c:idx"), "val"), i) : i;
    const pts: Record<number, RGBA> = {};
    for (const dp of findAll(s, "c:dPt")) {
      const c = color(find(dp, "c:spPr"));
      if (c) pts[int(attr(find(dp, "c:idx"), "val"))] = c;
    }
    return {
      name, cats: cache(catRef), vals: cache(find(s, "c:val/c:numRef")).map(toNum),
      color: color(find(s, "c:spPr")) ?? (accents.length ? accents[idx % accents.length] : [68, 114, 196, 1]),
      pts, code: textOf(find(s, "c:val/c:numRef/c:numCache/c:formatCode")) || "General",
      labels: labelsOn(s, pl), marker: markerOf(s, pl),
    };
  });
  if (!series.length) return "";
  const legend = find(chart, "c:chart/c:legend");
  const legPos = legend ? attr(find(legend, "c:legendPos"), "val") ?? "r" : null;
  const pie = kind === "pieChart" || kind === "doughnutChart";
  const items: [string, RGBA][] = pie
    ? series[0].cats.map((c, i) => [c, series[0].pts[i] ?? accents[i % accents.length]])
    : series.map((s) => [s.name, s.color]);
  let bottom = h - 8, right = w - 8, left = 8;
  if (legPos === "b" || legPos === "t") {
    const lh = fs * 1.8;
    const total = items.reduce((a, [t]) => a + fs * 1.5 + t.length * fs * 0.55 + fs, 0);
    let x = (w - total) / 2;
    const y = legPos === "b" ? bottom - lh / 2 : top + lh / 2;
    for (const [t, c] of items) {
      out.push(`<rect x="${f1(x)}" y="${f1(y - fs * 0.4)}" width="${f1(fs * 0.8)}" height="${f1(fs * 0.8)}" fill="${paint(c)}"/>`);
      out.push(`<text x="${f1(x + fs * 1.1)}" y="${f1(y + fs * 0.35)}" font-size="${f1(fs)}" fill="#595959" style="${font}">${esc(t)}</text>`);
      x += fs * 1.5 + t.length * fs * 0.55 + fs;
    }
    if (legPos === "b") bottom -= lh; else top += lh;
  } else if (legPos) {
    const lw = Math.max(...items.map(([t]) => t.length)) * fs * 0.55 + fs * 2;
    const x = legPos !== "l" ? right - lw + fs * 0.5 : left;
    let y = (top + bottom) / 2 - (items.length * fs * 1.5) / 2;
    for (const [t, c] of items) {
      out.push(`<rect x="${f1(x)}" y="${f1(y)}" width="${f1(fs * 0.8)}" height="${f1(fs * 0.8)}" fill="${paint(c)}"/>`);
      out.push(`<text x="${f1(x + fs * 1.1)}" y="${f1(y + fs * 0.75)}" font-size="${f1(fs)}" fill="#595959" style="${font}">${esc(t)}</text>`);
      y += fs * 1.5;
    }
    if (legPos === "l") left += lw; else right -= lw;
  }
  out.push(pie ? pieSvg(series[0], kind, pl, left, top, right, bottom, accents, fs, font)
               : xySvg(series, kind, pl, plot, left, top, right, bottom, fs, font));
  return `<svg width="${f1(w)}" height="${f1(h)}" viewBox="0 0 ${f1(w)} ${f1(h)}" style="position:absolute;left:0;top:0;overflow:visible">${out.join("")}</svg>`;
}

function bar(horizontal: boolean, c: number, bw: number, a: number, b: number, col: RGBA): string {
  const lo = Math.min(a, b), hi = Math.max(a, b);
  return horizontal
    ? `<rect x="${f1(lo)}" y="${f1(c)}" width="${f1(hi - lo)}" height="${f1(bw)}" fill="${paint(col)}"/>`
    : `<rect x="${f1(c)}" y="${f1(lo)}" width="${f1(bw)}" height="${f1(hi - lo)}" fill="${paint(col)}"/>`;
}

function label(horizontal: boolean, c: number, v: number, text: string, fs: number, font: string, inside: boolean): string {
  if (horizontal)
    return `<text x="${f1(v + (inside ? 0 : fs * 0.3))}" y="${f1(c + fs * 0.35)}" text-anchor="${inside ? "middle" : "start"}" font-size="${f1(fs)}" fill="#404040" style="${font}">${esc(text)}</text>`;
  return `<text x="${f1(c)}" y="${f1(v + (inside ? fs * 0.35 : -fs * 0.4))}" text-anchor="middle" font-size="${f1(fs)}" fill="#404040" style="${font}">${esc(text)}</text>`;
}

function xySvg(series: Series[], kind: string, pl: El, plot: El, left: number, top: number, right: number, bottom: number,
               fs: number, font: string): string {
  const out: string[] = [];
  const grouping = attr(find(pl, "c:grouping"), "val") ?? "clustered";
  const horizontal = kind === "barChart" && attr(find(pl, "c:barDir"), "val") === "bar";
  const cats = series.reduce((a, s) => (s.cats.length > a.length ? s.cats : a), [] as string[]);
  const n = Math.max(cats.length, ...series.map((s) => s.vals.length));
  const stacked = grouping === "stacked" || grouping === "percentStacked";
  const pct = grouping === "percentStacked";
  const vAt = (s: Series, i: number) => (i < s.vals.length ? s.vals[i] ?? 0 : 0);
  let lo: number, hi: number;
  if (stacked) {
    const pos = Array.from({ length: n }, (_, i) => series.reduce((a, s) => a + Math.max(0, vAt(s, i)), 0));
    const neg = Array.from({ length: n }, (_, i) => series.reduce((a, s) => a + Math.min(0, vAt(s, i)), 0));
    [lo, hi] = pct ? [0, 1] : [Math.min(0, ...neg), Math.max(0, ...pos)];
  } else {
    const all = series.flatMap((s) => s.vals.filter((v): v is number => v !== null));
    lo = Math.min(0, ...all); hi = Math.max(0, ...all);
  }
  let step: number;
  [lo, hi, step] = niceAxis(lo, hi);
  const sc = find(plot, "c:valAx/c:scaling");
  if (find(sc, "c:min")) lo = parseFloat(attr(find(sc, "c:min"), "val") ?? "0");
  if (find(sc, "c:max")) hi = parseFloat(attr(find(sc, "c:max"), "val") ?? "1");
  const code = series[0].code;
  const vax = find(plot, "c:valAx");
  const deleted = attr(find(vax, "c:delete"), "val") === "1";
  const grid = !!find(vax, "c:majorGridlines");
  const fmtAx = (v: number) => fmtValue(v, pct ? "0%" : code);
  const labelW = deleted ? 0 : Math.max(fmtAx(lo).length, fmtAx(hi).length) * fs * 0.6 + fs * 0.6;
  const catH = fs * 1.8;
  const x0 = left + (horizontal ? Math.max(...cats.map((c) => c.length)) * fs * 0.55 + fs : labelW);
  const x1 = right, y0 = top;
  const y1 = horizontal ? bottom - (deleted ? 0 : fs * 1.8) : bottom - catH;
  const vpos = (v: number) => { const u = (v - lo) / (hi - lo || 1); return horizontal ? x0 + u * (x1 - x0) : y1 - u * (y1 - y0); };
  for (let k = lo; k <= hi + step * 1e-6; k += step) {
    const p = vpos(k);
    if (grid) out.push(horizontal
      ? `<line x1="${f1(p)}" y1="${f1(y0)}" x2="${f1(p)}" y2="${f1(y1)}" stroke="#D9D9D9" stroke-width="1"/>`
      : `<line x1="${f1(x0)}" y1="${f1(p)}" x2="${f1(x1)}" y2="${f1(p)}" stroke="#D9D9D9" stroke-width="1"/>`);
    if (!deleted) out.push(horizontal
      ? `<text x="${f1(p)}" y="${f1(y1 + fs * 1.3)}" text-anchor="middle" font-size="${f1(fs)}" fill="#595959" style="${font}">${fmtAx(k)}</text>`
      : `<text x="${f1(x0 - fs * 0.4)}" y="${f1(p + fs * 0.35)}" text-anchor="end" font-size="${f1(fs)}" fill="#595959" style="${font}">${fmtAx(k)}</text>`);
  }
  const band = (horizontal ? y1 - y0 : x1 - x0) / Math.max(1, n);
  // a bar chart's category axis runs bottom-up in PowerPoint (first category lowest)
  const cpos = (i: number) => (horizontal ? y0 + (n - 1 - i) * band : x0 + i * band);
  for (let i = 0; i < n; i++) {
    const c0 = cpos(i), lab = esc(cats[i] ?? "");
    out.push(horizontal
      ? `<text x="${f1(x0 - fs * 0.4)}" y="${f1(c0 + band / 2 + fs * 0.35)}" text-anchor="end" font-size="${f1(fs)}" fill="#595959" style="${font}">${lab}</text>`
      : `<text x="${f1(c0 + band / 2)}" y="${f1(y1 + fs * 1.3)}" text-anchor="middle" font-size="${f1(fs)}" fill="#595959" style="${font}">${lab}</text>`);
  }
  const base = vpos(Math.max(lo, Math.min(0, hi)));
  if (kind === "barChart") {
    const gap = int(attr(find(pl, "c:gapWidth"), "val"), 150) / 100;
    const ov = (find(pl, "c:overlap") ? int(attr(find(pl, "c:overlap"), "val")) : stacked ? 100 : 0) / 100;
    const ns = series.length;
    for (let i = 0; i < n; i++) {
      const c0 = cpos(i);
      if (stacked) {
        const bw = band / (1 + gap);
        let accP = 0, accN = 0;
        const tot = series.reduce((a, s) => a + Math.abs(vAt(s, i)), 0) || 1;
        for (const s of series) {
          let v = vAt(s, i);
          if (pct) v /= tot;
          const a0 = v >= 0 ? accP : accN, a1 = a0 + v;
          if (v >= 0) accP = a1; else accN = a1;
          out.push(bar(horizontal, c0 + (band - bw) / 2, bw, vpos(a0), vpos(a1), s.pts[i] ?? s.color));
          if (s.labels) out.push(label(horizontal, c0 + band / 2, (vpos(a0) + vpos(a1)) / 2, fmtValue(v, pct ? "0%" : s.code), fs, font, true));
        }
      } else {
        const bw = band / (ns - ov * (ns - 1) + gap);
        const start = c0 + (band - bw * (ns - ov * (ns - 1))) / 2;
        series.forEach((s, j) => {
          const v = i < s.vals.length ? s.vals[i] : null;
          if (v === null) return;
          const bx = start + j * bw * (1 - ov);
          out.push(bar(horizontal, bx, bw, base, vpos(v), s.pts[i] ?? s.color));
          if (s.labels) out.push(label(horizontal, bx + bw / 2, vpos(v), fmtValue(v, s.code), fs, font, false));
        });
      }
    }
  } else {
    const acc = new Array(n).fill(0);
    for (const s of series) {
      const pts: [number, number, number][] = [];
      for (let i = 0; i < n; i++) {
        let v = i < s.vals.length ? s.vals[i] : null;
        if (v === null) continue;
        if (stacked) { acc[i] += v; v = acc[i]; }
        pts.push([x0 + (i + 0.5) * band, vpos(v as number), v as number]);
      }
      if (!pts.length) continue;
      const col = paint(s.color);
      if (kind === "areaChart") {
        out.push(`<path d="M${f1(pts[0][0])} ${f1(base)} ${pts.map(([x, y]) => `L${f1(x)} ${f1(y)}`).join(" ")} L${f1(pts[pts.length - 1][0])} ${f1(base)} Z" fill="${col}"/>`);
      } else {
        out.push(`<path d="M${pts.map(([x, y]) => `${f1(x)} ${f1(y)}`).join(" L")}" fill="none" stroke="${col}" stroke-width="${f1(fs * 0.22)}" stroke-linejoin="round" stroke-linecap="round"/>`);
        if (s.marker) {
          const r = fs * 0.3;
          for (const [x, y] of pts) out.push(s.marker === "square"
            ? `<rect x="${f1(x - r)}" y="${f1(y - r)}" width="${f1(2 * r)}" height="${f1(2 * r)}" fill="${col}"/>`
            : `<circle cx="${f1(x)}" cy="${f1(y)}" r="${f1(r)}" fill="${col}"/>`);
        }
      }
      if (s.labels) for (const [x, y, v] of pts) out.push(label(false, x, y, fmtValue(v, s.code), fs, font, false));
    }
  }
  out.push(horizontal
    ? `<line x1="${f1(base)}" y1="${f1(y0)}" x2="${f1(base)}" y2="${f1(y1)}" stroke="#BFBFBF" stroke-width="1"/>`
    : `<line x1="${f1(x0)}" y1="${f1(base)}" x2="${f1(x1)}" y2="${f1(base)}" stroke="#BFBFBF" stroke-width="1"/>`);
  return out.join("");
}

function pieSvg(s: Series, kind: string, pl: El, left: number, top: number, right: number, bottom: number,
                accents: RGBA[], fs: number, font: string): string {
  const vals = s.vals.map((v) => Math.max(0, v ?? 0));
  const tot = vals.reduce((a, b) => a + b, 0) || 1;
  const cx = (left + right) / 2, cy = (top + bottom) / 2;
  const r = (Math.min(right - left, bottom - top) / 2) * 0.92;
  const hole = kind === "doughnutChart" ? int(attr(find(pl, "c:holeSize"), "val"), 50) / 100 : 0;
  let ang = ((int(attr(find(pl, "c:firstSliceAng"), "val")) - 90) * Math.PI) / 180;
  const out: string[] = [];
  vals.forEach((v, i) => {
    const a1 = ang + (v / tot) * 2 * Math.PI;
    const col = s.pts[i] ?? (accents.length ? accents[i % accents.length] : s.color);
    const large = a1 - ang > Math.PI ? 1 : 0;
    const p0 = [cx + r * Math.cos(ang), cy + r * Math.sin(ang)], p1 = [cx + r * Math.cos(a1), cy + r * Math.sin(a1)];
    let d: string;
    if (v >= tot - 1e-9) d = `M${f1(cx - r)} ${f1(cy)} A${f1(r)} ${f1(r)} 0 1 1 ${f1(cx + r)} ${f1(cy)} A${f1(r)} ${f1(r)} 0 1 1 ${f1(cx - r)} ${f1(cy)} Z`;
    else if (hole) {
      const ri = r * hole;
      const q0 = [cx + ri * Math.cos(a1), cy + ri * Math.sin(a1)], q1 = [cx + ri * Math.cos(ang), cy + ri * Math.sin(ang)];
      d = `M${f1(p0[0])} ${f1(p0[1])} A${f1(r)} ${f1(r)} 0 ${large} 1 ${f1(p1[0])} ${f1(p1[1])} L${f1(q0[0])} ${f1(q0[1])} A${f1(ri)} ${f1(ri)} 0 ${large} 0 ${f1(q1[0])} ${f1(q1[1])} Z`;
    } else d = `M${f1(cx)} ${f1(cy)} L${f1(p0[0])} ${f1(p0[1])} A${f1(r)} ${f1(r)} 0 ${large} 1 ${f1(p1[0])} ${f1(p1[1])} Z`;
    out.push(`<path d="${d}" fill="${paint(col)}" stroke="#FFFFFF" stroke-width="1.5"/>`);
    if (s.labels && v > 0) {
      const mid = (ang + a1) / 2, rr = r * (hole ? (1 + hole) / 2 : 0.7);
      out.push(`<text x="${f1(cx + rr * Math.cos(mid))}" y="${f1(cy + rr * Math.sin(mid) + fs * 0.35)}" text-anchor="middle" font-size="${f1(fs)}" fill="#FFFFFF" style="${font}">${esc(fmtValue(v, s.code))}</text>`);
    }
    ang = a1;
  });
  return out.join("");
}
