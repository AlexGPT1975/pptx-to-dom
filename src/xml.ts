// Small namespace-aware XML helpers over the standard DOM API (@xmldom/xmldom in Node, the
// same interface in browsers). Paths use prefixes from NS: "a:xfrm/a:off", ".//a:t" for any
// descendant, and one attribute predicate "a:ext[@uri='...']".
import { DOMParser } from "@xmldom/xmldom";

export const NS: Record<string, string> = {
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  dsp: "http://schemas.microsoft.com/office/drawing/2008/diagram",
  asvg: "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
};

export type El = Element;

export function parseXml(text: string): El {
  const doc = new DOMParser().parseFromString(text, "text/xml");
  return doc.documentElement as unknown as El;
}

export const ln = (el: El): string => el.localName || el.nodeName.replace(/^.*:/, "");

export function kids(el: El | null | undefined): El[] {
  if (!el) return [];
  const out: El[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as El);
  return out;
}

/** First child with this local name, whatever its namespace (p:spPr / dsp:spPr). */
export function kid(el: El | null | undefined, name: string): El | null {
  for (const c of kids(el)) if (ln(c) === name) return c;
  return null;
}

function stepMatch(el: El, step: string): boolean {
  const m = step.match(/^(?:(\w+):)?([\w*]+)(?:\[@(\w+)='([^']*)'\])?$/);
  if (!m) return false;
  const [, pre, local, attr, val] = m;
  if (local !== "*" && ln(el) !== local) return false;
  if (pre && NS[pre] && el.namespaceURI !== NS[pre]) return false;
  if (attr && el.getAttribute(attr) !== val) return false;
  return true;
}

function descendants(el: El): El[] {
  const out: El[] = [];
  const walk = (e: El) => { for (const c of kids(e)) { out.push(c); walk(c); } };
  walk(el);
  return out;
}

/** All elements matching a path relative to el. */
export function findAll(el: El | null | undefined, path: string): El[] {
  if (!el) return [];
  let cur: El[] = [el];
  let p = path;
  let deep = false;
  if (p.startsWith(".//")) { deep = true; p = p.slice(3); } else if (p.startsWith("./")) p = p.slice(2);
  const steps = p.split("/").filter(Boolean);
  steps.forEach((step, i) => {
    const next: El[] = [];
    for (const e of cur) {
      const pool = i === 0 && deep ? descendants(e) : kids(e);
      for (const c of pool) if (stepMatch(c, step)) next.push(c);
    }
    cur = next;
  });
  return cur;
}

export const find = (el: El | null | undefined, path: string): El | null => findAll(el, path)[0] ?? null;

export const attr = (el: El | null | undefined, name: string): string | null =>
  el && el.hasAttribute(name) ? el.getAttribute(name) : null;

export function attrs(el: El | null | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!el) return out;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    out[a.localName || a.name] = a.value;
  }
  return out;
}

export function int(v: string | null | undefined, dflt = 0): number {
  if (v === null || v === undefined || v === "") return dflt;
  const n = parseInt(v, 10);
  if (!isNaN(n)) return n;
  const f = parseFloat(v);
  return isNaN(f) ? dflt : Math.trunc(f);
}

export const rId = (el: El | null | undefined, name = "embed"): string | null =>
  el ? (el.getAttributeNS(NS.r, name) || el.getAttribute(`r:${name}`)) : null;

export function textOf(el: El | null | undefined): string {
  return el ? (el.textContent ?? "") : "";
}
