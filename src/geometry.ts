// DrawingML geometry -> SVG paths: every PowerPoint preset shape, plus custom geometry.
//
// PowerPoint stores a preset shape only by name (<a:prstGeom prst="chevron">) and its adjust
// values (<a:avLst>); the outline is defined in ECMA-376's presetShapeDefinitions as a small
// formula language (guides) and path commands. Evaluating it gives the exact outline of all
// 187 presets at any size and adjustment. Custom geometry (<a:custGeom>) uses the same language.
import { PRESETS, PresetDef } from "./presets.generated.js";
import { El, attr, find, findAll, kids, ln } from "./xml.js";

const ANG = 60000; // DrawingML angles: 1/60000 degree

export interface ShapePath { d: string; fill: boolean; stroke: boolean; shade: string | null }
export interface ShapeGeometry { paths: ShapePath[]; text: [number, number, number, number] }

export const hasPreset = (name: string): boolean => name in PRESETS;

function builtins(w: number, h: number): Record<string, number> {
  const ss = Math.min(w, h), ls = Math.max(w, h);
  const v: Record<string, number> = {
    l: 0, t: 0, r: w, b: h, w, h, hc: w / 2, vc: h / 2, ss, ls,
    cd2: 10800000, cd4: 5400000, cd8: 2700000, "3cd4": 16200000, "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000,
  };
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) { v[`wd${n}`] = w / n; v[`hd${n}`] = h / n; v[`ssd${n}`] = ss / n; }
  return v;
}

const num = (tok: string, env: Record<string, number>): number => {
  if (tok in env) return env[tok];
  const f = parseFloat(tok);
  return isNaN(f) ? 0 : f;
};
const rad = (a: number) => (a / ANG) * Math.PI / 180;

export function evalFormula(fmla: string, env: Record<string, number>): number {
  const p = fmla.trim().split(/\s+/);
  const op = p[0];
  const a = p.slice(1).map((t) => num(t, env));
  const [x = 0, y = 0, z = 0] = a;
  switch (op) {
    case "val": return x;
    case "*/": return z ? (x * y) / z : 0;
    case "+-": return x + y - z;
    case "+/": return z ? (x + y) / z : 0;
    case "?:": return x > 0 ? y : z;
    case "abs": return Math.abs(x);
    case "at2": return (Math.atan2(y, x) * 180 / Math.PI) * ANG;
    case "cat2": return x * Math.cos(Math.atan2(z, y));
    case "sat2": return x * Math.sin(Math.atan2(z, y));
    case "cos": return x * Math.cos(rad(y));
    case "sin": return x * Math.sin(rad(y));
    case "tan": return x * Math.tan(rad(y));
    case "max": return Math.max(x, y);
    case "min": return Math.min(x, y);
    case "mod": return Math.sqrt(x * x + y * y + z * z);
    case "pin": return y < x ? x : y > z ? z : y;
    case "sqrt": return Math.sqrt(Math.max(0, x));
    default: return 0;
  }
}

const fmt = (v: number): string => {
  const s = (Math.round(v * 100) / 100).toString();
  return s === "-0" ? "0" : s;
};

interface PathSpec { w: string; h: string; fill: string; stroke: string; c: string[][] }

function pathD(path: PathSpec, env: Record<string, number>, w: number, h: number): string {
  const pw = parseFloat(path.w || "0") || 0, ph = parseFloat(path.h || "0") || 0;
  const sx = pw ? w / pw : 1, sy = ph ? h / ph : 1;
  const d: string[] = [];
  let cx = 0, cy = 0;
  for (const cmd of path.c) {
    const k = cmd[0];
    if (k === "M" || k === "L") {
      cx = num(cmd[1], env) * sx; cy = num(cmd[2], env) * sy;
      d.push(`${k}${fmt(cx)} ${fmt(cy)}`);
    } else if (k === "C" && cmd.length >= 7) {
      const p = [1, 2, 3, 4, 5, 6].map((i) => num(cmd[i], env) * (i % 2 ? sx : sy));
      cx = p[4]; cy = p[5];
      d.push(`C${p.map(fmt).join(" ")}`);
    } else if (k === "Q" && cmd.length >= 5) {
      const p = [1, 2, 3, 4].map((i) => num(cmd[i], env) * (i % 2 ? sx : sy));
      cx = p[2]; cy = p[3];
      d.push(`Q${p.map(fmt).join(" ")}`);
    } else if (k === "A") {
      const wr = num(cmd[1], env) * sx, hr = num(cmd[2], env) * sy;
      const st = num(cmd[3], env), sw = num(cmd[4], env);
      if (wr <= 0 || hr <= 0 || sw === 0) continue;
      // DrawingML angles are VISUAL angles on the ellipse; convert to the parametric angle
      const par = (deg: number) => { const a = deg * Math.PI / 180; return Math.atan2(wr * Math.sin(a), hr * Math.cos(a)); };
      const a0 = par(st / ANG);
      let a1 = par((st + sw) / ANG);
      const sweepDeg = sw / ANG;
      const turns = Math.floor(Math.abs(sweepDeg) / 360);
      if (sweepDeg > 0) { while (a1 <= a0) a1 += 2 * Math.PI; a1 += 2 * Math.PI * turns; }
      else { while (a1 >= a0) a1 -= 2 * Math.PI; a1 -= 2 * Math.PI * turns; }
      const ox = cx - wr * Math.cos(a0), oy = cy - hr * Math.sin(a0);
      const total = a1 - a0;
      const n = Math.max(1, Math.ceil(Math.abs(total) / Math.PI - 1e-9)); // SVG arcs: <=180° pieces
      for (let i = 1; i <= n; i++) {
        const t = a0 + (total * i) / n;
        const ex = ox + wr * Math.cos(t), ey = oy + hr * Math.sin(t);
        d.push(`A${fmt(wr)} ${fmt(hr)} 0 ${Math.abs(total / n) > Math.PI ? 1 : 0} ${total > 0 ? 1 : 0} ${fmt(ex)} ${fmt(ey)}`);
        cx = ex; cy = ey;
      }
    } else if (k === "Z") {
      d.push("Z");
    }
  }
  return d.join(" ");
}

function avOverrides(geom: El | null): Record<string, string> {
  const out: Record<string, string> = {};
  for (const gd of findAll(geom, "a:avLst/a:gd")) out[attr(gd, "name") ?? ""] = attr(gd, "fmla") ?? "val 0";
  return out;
}

function guides(def: PresetDef, overrides: Record<string, string>, env: Record<string, number>): void {
  for (const [name, fmla] of def.av) env[name] = evalFormula(overrides[name] ?? fmla, env);
  for (const [name, fmla] of Object.entries(overrides)) if (!(name in env)) env[name] = evalFormula(fmla, env);
  for (const [name, fmla] of def.gd) env[name] = evalFormula(fmla, env);
}

/** A custom geometry element in the same shape as a preset definition. */
function custDef(geom: El): PresetDef {
  const def: PresetDef = { av: [], gd: [], rect: null, paths: [] };
  for (const gd of findAll(geom, "a:avLst/a:gd")) def.av.push([attr(gd, "name") ?? "", attr(gd, "fmla") ?? "val 0"]);
  for (const gd of findAll(geom, "a:gdLst/a:gd")) def.gd.push([attr(gd, "name") ?? "", attr(gd, "fmla") ?? "val 0"]);
  const r = find(geom, "a:rect");
  if (r) def.rect = ["l", "t", "r", "b"].map((k) => attr(r, k) ?? k);
  for (const p of findAll(geom, "a:pathLst/a:path")) {
    const spec: PathSpec = { w: attr(p, "w") ?? "", h: attr(p, "h") ?? "", fill: attr(p, "fill") ?? "norm",
                             stroke: attr(p, "stroke") ?? "true", c: [] };
    for (const cmd of kids(p)) {
      const t = ln(cmd);
      const pts = findAll(cmd, "a:pt").flatMap((pt) => [attr(pt, "x") ?? "0", attr(pt, "y") ?? "0"]);
      if (t === "arcTo") spec.c.push(["A", attr(cmd, "wR") ?? "0", attr(cmd, "hR") ?? "0", attr(cmd, "stAng") ?? "0", attr(cmd, "swAng") ?? "0"]);
      else if (t === "close") spec.c.push(["Z"]);
      else spec.c.push([({ moveTo: "M", lnTo: "L", cubicBezTo: "C", quadBezTo: "Q" } as Record<string, string>)[t] ?? "", ...pts]);
    }
    def.paths.push(spec);
  }
  return def;
}

/** <a:prstGeom> or <a:custGeom> at w x h px -> paths + text rectangle, or null. */
export function shapePaths(geom: El | null, w: number, h: number): ShapeGeometry | null {
  if (!geom) return null;
  const env = builtins(w, h);
  let def: PresetDef | undefined;
  if (ln(geom) === "prstGeom") {
    def = PRESETS[attr(geom, "prst") ?? "rect"];
    if (!def) return null;
    guides(def, avOverrides(geom), env);
  } else if (ln(geom) === "custGeom") {
    def = custDef(geom);
    guides(def, {}, env);
  } else return null;
  if (!def.paths.length) return null;
  const paths: ShapePath[] = [];
  for (const p of def.paths) {
    const d = pathD(p, env, w, h);
    if (!d) continue;
    paths.push({ d, fill: p.fill !== "none", shade: p.fill !== "norm" && p.fill !== "none" ? p.fill : null,
                 stroke: p.stroke !== "false" && p.stroke !== "0" });
  }
  const r = def.rect;
  const text: [number, number, number, number] = r
    ? [num(r[0] ?? "l", env), num(r[1] ?? "t", env), num(r[2] ?? "r", env), num(r[3] ?? "b", env)]
    : [0, 0, w, h];
  return { paths, text };
}

/** One adjust value of a preset (raw DrawingML units), e.g. roundRect's "adj". */
export function adjust(geom: El | null, name: string, dflt: number): number {
  const f = avOverrides(geom)[name];
  if (!f) return dflt;
  const v = parseFloat(f.trim().split(/\s+/).pop() ?? "");
  return isNaN(v) ? dflt : v;
}

/** Bounding box of an SVG path from shapePaths (M/L/C/Q/A/Z, absolute). */
export function pathBBox(d: string): [number, number, number, number] | null {
  const xs: number[] = [], ys: number[] = [];
  const re = /([MLCQAZ])([^MLCQAZ]*)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(d))) {
    const v = m[2].trim().split(/\s+/).filter(Boolean).map(Number);
    if (m[1] === "A") {
      for (let i = 0; i + 6 < v.length; i += 7) {   // an arc bulges up to its radius past its end

        xs.push(v[i + 5], v[i + 5] - v[i], v[i + 5] + v[i]);
        ys.push(v[i + 6], v[i + 6] - v[i + 1], v[i + 6] + v[i + 1]);
      }
    } else {
      for (let i = 0; i + 1 < v.length; i += 2) { xs.push(v[i]); ys.push(v[i + 1]); }
    }
  }
  if (!xs.length) return null;
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}
