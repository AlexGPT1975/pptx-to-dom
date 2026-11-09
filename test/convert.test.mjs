// Unit tests on small synthetic .pptx files built in memory (no binary fixtures).
// Run: npm test  (builds first, then node --test)
import assert from "node:assert/strict";
import { test } from "node:test";
import JSZip from "jszip";
import { PptxImportError, adjust, convertPptx, shapePaths } from "../dist/index.js";
import { applyMods, css, linearLight, patternSvg } from "../dist/color.js";
import { fmtValue, niceAxis } from "../dist/charts.js";
import { parseXml } from "../dist/xml.js";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NSDECL = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

async function pptx(body, extra = {}) {
  const z = new JSZip();
  z.file("ppt/presentation.xml", `<p:presentation ${NSDECL}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
  z.file("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  z.file("ppt/slides/slide1.xml", `<p:sld ${NSDECL}><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`);
  for (const [k, v] of Object.entries(extra)) z.file(k, v);
  return z.generateAsync({ type: "uint8array" });
}

const shape = (geom, extra = "", txt = "") =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
  '<a:xfrm><a:off x="635000" y="635000"/><a:ext cx="2540000" cy="1270000"/></a:xfrm>' +
  `<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EE0000"/></a:solidFill>${extra}</p:spPr>` +
  (txt ? `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>${txt}</a:t></a:r></a:p></p:txBody>` : "") + "</p:sp>";

const geom = (prst, adj = {}) => parseXml(`<a:prstGeom xmlns:a="${A}" prst="${prst}"><a:avLst>${Object.entries(adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${v}"/>`).join("")}</a:avLst></a:prstGeom>`);

test("chevron point follows its adjust value", () => {
  const g = shapePaths(geom("chevron", { adj: 30000 }), 360, 130);
  assert.equal(g.paths[0].d, "M0 0 L321 0 L360 65 L321 130 L0 130 L39 65 Z");
  assert.deepEqual(g.text, [39, 0, 321, 130]);
});

test("round rectangle insets its text by the corner", () => {
  const [l, , r] = shapePaths(geom("roundRect", { adj: 10000 }), 320, 170).text;
  assert.ok(l > 4.5 && l < 5.5 && r > 314.5 && r < 315.5);
  assert.equal(adjust(geom("roundRect", { adj: 10000 }), "adj", 16667), 10000);
});

test("block arc sweeps three quarters", () => {
  const d = shapePaths(geom("blockArc", { adj1: 16200000, adj2: 10800000, adj3: 12000 }), 280, 280).paths[0].d;
  assert.ok(d.startsWith("M140 0 A140 140") && d.includes("L33.6 140"));
});

test("colour modifiers", () => {
  const c = applyMods([238, 0, 0, 1], parseXml(`<a:srgbClr xmlns:a="${A}" val="EE0000"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:srgbClr>`));
  assert.equal(css(c), "#FFC9C9");
  assert.equal(css(applyMods([255, 255, 255, 1], parseXml(`<a:srgbClr xmlns:a="${A}" val="FFFFFF"><a:alpha val="14000"/></a:srgbClr>`))), "rgba(255,255,255,0.140)");
});

test("gradients blend in linear light", () => {
  const stops = linearLight([[0, [238, 0, 0, 1]], [1, [255, 138, 138, 1]]]);
  const mid = stops.reduce((a, s) => (Math.abs(s[0] - 0.5) < Math.abs(a[0] - 0.5) ? s : a));
  assert.ok(mid[1][1] > 90); // sRGB blending would give ~69
});

test("pattern fills become tiles", () => {
  assert.ok(patternSvg("dkDnDiag", [238, 0, 0, 1], [255, 255, 255, 1], 1.5).startsWith("data:image/svg+xml;base64,"));
  assert.equal(patternSvg("nope", [0, 0, 0, 1], [255, 255, 255, 1], 1.5), null);
});

test("chart axis and number format like PowerPoint", () => {
  assert.deepEqual(niceAxis(0, 23.9), [0, 30, 5]);
  assert.deepEqual(niceAxis(0, 60), [0, 70, 10]);
  assert.equal(fmtValue(1234.5, "#,##0.0"), "1.234,5");
  assert.equal(fmtValue(0.62, "0%"), "62 %");
});

test("preset shapes become svg, rectangles stay css boxes", async () => {
  const r = await convertPptx(await pptx(shape("chevron") + shape("rect")));
  const html = r.slides[0].html;
  assert.ok(html.includes("<svg") && html.includes('d="M0 0 L'));
  assert.equal(html.split("background:#EE0000").length - 1, 1);
});

test("shadow becomes box-shadow", async () => {
  const eff = '<a:effectLst><a:outerShdw blurRad="190500" dist="63500" dir="5400000"><a:srgbClr val="000000"><a:alpha val="20000"/></a:srgbClr></a:outerShdw></a:effectLst>';
  const r = await convertPptx(await pptx(shape("roundRect", eff)));
  assert.match(r.slides[0].html, /box-shadow:0\.0px 10\.0px 30\.0px rgba\(0,0,0,0\.200\)/);
});

test("text is never clipped; WordArt is flagged", async () => {
  const body = shape("rect", "", "Bogen").replace("<p:txBody><a:bodyPr/>", '<p:txBody><a:bodyPr><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></a:bodyPr>');
  const r = await convertPptx(await pptx(body));
  assert.ok(r.slides[0].html.includes("overflow:visible") && !r.slides[0].html.includes("overflow:hidden"));
  assert.deepEqual(r.slides[0].flags, ["wordart"]);
  assert.equal(r.title, "Bogen");
});

test("default insets are PowerPoint's 0.1 inch", async () => {
  const r = await convertPptx(await pptx(shape("rect", "", "Text")));
  const m = r.slides[0].html.match(/class="pp-text" style="position:absolute;left:([\d.]+)px/);
  assert.equal(parseFloat(m[1]), 14.4);
});

test("garbage input raises PptxImportError", async () => {
  await assert.rejects(convertPptx(new Uint8Array([1, 2, 3])), PptxImportError);
  const z = new JSZip(); z.file("x.txt", "x");
  await assert.rejects(convertPptx(await z.generateAsync({ type: "uint8array" })), PptxImportError);
});
