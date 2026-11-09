#!/usr/bin/env node
// pptx-to-dom deck.pptx [out.html] - writes a standalone HTML document with every slide.
import { readFileSync, writeFileSync } from "node:fs";
import { convertPptx, toHtmlDocument } from "./index.js";

const [, , input, output] = process.argv;
if (!input) {
  console.error("Aufruf: pptx-to-dom datei.pptx [ziel.html]");
  process.exit(2);
}
const result = await convertPptx(new Uint8Array(readFileSync(input)));
const target = output ?? input.replace(/\.pptx?$/i, "") + ".html";
writeFileSync(target, toHtmlDocument(result));
const flagged = result.slides.filter((s) => s.flags.length).map((s) => `${s.number}: ${s.flags.join(",")}`);
console.log(`${target}: ${result.slides.length} Folien, Schriften: ${result.fontsUsed.join(", ")}` +
            (flagged.length ? `\nnur angenähert: ${flagged.join("; ")}` : ""));
