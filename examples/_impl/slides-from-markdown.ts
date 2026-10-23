// SPEC-106: slide Markdown in, a deck out. Reads the Markdown file named on
// the command line (or a built-in sample), parses it with
// parseSlideMarkdownWithReport — printing everything the slide grammar does
// not carry — and writes the deck in the format its front matter names
// (pptx by default) through SPEC-080's writers. Run via
// `node examples/slides-from-markdown.js [deck.md]`; writes
// examples/output/<title>.pptx|.odp.
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { parseSlideMarkdownWithReport } from "../../src/lib/generate/slideMarkdown";
import { odpFileName, writeOdp } from "../../src/lib/presentation/odp/writer";
import { pptxFileName, writePptx } from "../../src/lib/presentation/pptx/writer";

const SAMPLE = `---
title: Release notes
---
# Release notes
Version 4.2 — October 2026

---

## What's new
- **Faster** opening of large decks
  - Up to _three times_ quicker on 200-slide files
- Generated decks from Markdown

Notes: Demo the open time on the big deck.

---

## Fixed
1. Notes survive a copy as ODP
2. Bullets keep their level on paste

> Thanks to everyone who filed reports.
`;

// argv[2] is the Markdown file when run as `node examples/slides-from-markdown.js deck.md`.
const source = process.argv[2] ? await readFile(process.argv[2], "utf8") : SAMPLE;
const { deck, dropped } = parseSlideMarkdownWithReport(source);

for (const record of dropped) {
  console.log(`Dropped (${record.kind}, ${record.location}): ${record.reason}`);
}

const outDir = fileURLToPath(new URL("../output/", import.meta.url));
await mkdir(outDir, { recursive: true });
const { bytes } = deck.format === "odp" ? await writeOdp(deck) : await writePptx(deck);
const outPath = `${outDir}${deck.format === "odp" ? odpFileName(deck.title) : pptxFileName(deck.title)}`;
await writeFile(outPath, bytes);
console.log(`Wrote ${outPath} (${deck.slides.length} slides)`);
