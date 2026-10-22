#!/usr/bin/env node
// Deterministic generator for the SPEC-047 performance-benchmark fixture: a
// ~300-page .docx with 500 images (~100 KB average), 200 tables and 400
// footnotes, body text at the document default size (specs/047's "The
// benchmark fixture"). Not committed to git — e2e/SPEC-047.perf.spec.ts
// builds it on demand if e2e/fixtures/spec-047-fixture.docx is absent.
//
// Run manually with:  npx vite-node scripts/build-spec047-fixture.mjs
// (vite-node is already a transitive dependency of vitest — no new
// dependency, matching this repo's protected package.json.)
//
// Deterministic means: a fixed document id (not crypto.randomUUID(), which
// createDocument() defaults to), a fixed `modified` date passed to
// writeDocx (which otherwise defaults to `new Date()`), fixed note ids
// (not createNoteId(), same reason), and a seeded PRNG for image pixel
// data — so the same script produces the same bytes on every run, and a
// heap/timing number from one run is comparable to the next (AC-25).
import { writeFileSync, mkdirSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { writeDocx } from "../src/lib/docx/writer";
import { createDocument, createNoteRun } from "../src/lib/documentModel";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// --- deterministic PRNG (mulberry32) ----------------------------------------
function mulberry32(seed) {
  let a = seed >>> 0;
  return function next() {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// --- minimal PNG encoder (24-bit RGB, no filtering), same shape as
// scripts/build-spec024-fixtures.mjs's buildPng but fed pseudo-random pixel
// data: a solid colour deflates to a few hundred bytes, which would defeat
// the "~100 KB average" the fixture is meant to hold — near-random data is
// close to incompressible, so the encoded size tracks the pixel count. ----
function u32be(value) {
  const buf = Buffer.alloc(4);
  buf.writeUInt32BE(value >>> 0, 0);
  return buf;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const typeBuf = Buffer.from(type, "ascii");
  const body = Buffer.concat([typeBuf, data]);
  return Buffer.concat([u32be(data.length), body, u32be(crc32(body))]);
}

function buildNoisePng(width, height, seed) {
  const rng = mulberry32(seed);
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (rowBytes + 1);
    raw[rowStart] = 0; // filter type: none
    for (let x = 0; x < width; x++) {
      const px = rowStart + 1 + x * 3;
      raw[px] = Math.floor(rng() * 256);
      raw[px + 1] = Math.floor(rng() * 256);
      raw[px + 2] = Math.floor(rng() * 256);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type: RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

// --- fixture composition ----------------------------------------------------
// The fixture sits deliberately below the supported ceiling (25,000 blocks,
// 1,000 images, 500 tables, 1,000 notes — specs/047's "The supported
// ceiling"): the benchmark measures the product working, not refusing.
const IMAGE_COUNT = 500;
const TABLE_COUNT = 200;
const FOOTNOTE_COUNT = 400;
// Sized (with SPEC-010's own "~600 paragraphs ~= 200 pages" precedent,
// scaled up and reduced for the page space images/tables also consume) to
// land in the neighbourhood of 300 pages — the spec's numbers are a
// composition target, not a value any acceptance criterion asserts exactly.
const FILLER_PARAGRAPH_COUNT = 1800;
// 220x140 px of near-random RGB deflates close to its raw ~92.5 KB size
// (noise barely compresses) — close to the "~100 KB average" the spec asks
// for, while leaving headroom under SPEC-009's 50 MB file-size ceiling
// (which applies to every open path, this fixture included — 500 images at
// a literal 100 KB average alone lands within a few KB of that ceiling,
// leaving no room for the other 200 tables/400 footnotes/1800 paragraphs of
// XML this fixture also needs).
const IMAGE_WIDTH_PX = 220;
const IMAGE_HEIGHT_PX = 140;

const SENTENCE =
  "The quick brown fox jumps over the lazy dog while the committee reviews the quarterly report. ";

function fillerParagraph(index) {
  const repeats = 3 + (index % 3); // varies paragraph length a little
  return { type: "paragraph", runs: [{ text: SENTENCE.repeat(repeats), formatting: {} }] };
}

function tableBlock(index) {
  const rows = 4;
  const cols = 3;
  return {
    type: "table",
    columnWidthsCm: Array.from({ length: cols }, () => 5),
    rows: Array.from({ length: rows }, (_, r) => ({
      cells: Array.from({ length: cols }, (_, c) => ({
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: `Table ${index} cell R${r + 1}C${c + 1}`, formatting: {} }],
          },
        ],
      })),
    })),
  };
}

function build() {
  const model = createDocument("spec-047-fixture");
  model.format = "docx";
  model.title = "SPEC-047 performance fixture";
  model.media = [];
  model.notes = [];

  const blocks = [];
  let imagesPlaced = 0;
  let tablesPlaced = 0;
  let footnotesPlaced = 0;

  for (let i = 0; i < FILLER_PARAGRAPH_COUNT; i++) {
    const paragraph = fillerParagraph(i);

    // One footnote reference roughly every 4th paragraph, up to the target
    // count — a superscript reference run appended to the paragraph's own
    // text (SPEC-031's shape: see createNoteRun).
    if (footnotesPlaced < FOOTNOTE_COUNT && i % 4 === 0) {
      const noteId = `spec047-note-${footnotesPlaced}`;
      model.notes.push({
        noteId,
        kind: "footnote",
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: `Footnote content number ${footnotesPlaced}.`, formatting: {} }],
          },
        ],
      });
      paragraph.runs.push(createNoteRun(noteId, "footnote"));
      footnotesPlaced++;
    }

    blocks.push(paragraph);

    // One image roughly every 3.6 paragraphs, up to the target count.
    if (imagesPlaced < IMAGE_COUNT && i % 4 === 0) {
      const mediaId = `spec047-image-${imagesPlaced}`;
      const png = buildNoisePng(IMAGE_WIDTH_PX, IMAGE_HEIGHT_PX, imagesPlaced + 1);
      model.media.push({ mediaId, contentType: "image/png", data: png.toString("base64") });
      blocks.push({
        type: "image",
        mediaId,
        naturalWidthPx: IMAGE_WIDTH_PX,
        naturalHeightPx: IMAGE_HEIGHT_PX,
        widthCm: 6,
        heightCm: 3.6,
        wrap: "inline",
      });
      imagesPlaced++;
    }

    // One table roughly every 9th paragraph, up to the target count.
    if (tablesPlaced < TABLE_COUNT && i % 9 === 0) {
      blocks.push(tableBlock(tablesPlaced));
      tablesPlaced++;
    }
  }

  // Any remainder (rounding in the interleave above) goes at the end so the
  // counts always land exactly on the documented composition.
  while (imagesPlaced < IMAGE_COUNT) {
    const mediaId = `spec047-image-${imagesPlaced}`;
    const png = buildNoisePng(IMAGE_WIDTH_PX, IMAGE_HEIGHT_PX, imagesPlaced + 1);
    model.media.push({ mediaId, contentType: "image/png", data: png.toString("base64") });
    blocks.push({
      type: "image",
      mediaId,
      naturalWidthPx: IMAGE_WIDTH_PX,
      naturalHeightPx: IMAGE_HEIGHT_PX,
      widthCm: 6,
      heightCm: 3.6,
      wrap: "inline",
    });
    imagesPlaced++;
  }
  while (tablesPlaced < TABLE_COUNT) {
    blocks.push(tableBlock(tablesPlaced));
    tablesPlaced++;
  }

  // A trailing plain paragraph, always: the interleave above can end on a
  // table (its remainder loop runs last), and a caret can't sit "after" a
  // table the way it can after a paragraph — AC-6's "typing at the document
  // end" needs an ordinary paragraph there to type into.
  blocks.push({ type: "paragraph", runs: [{ text: "End of document.", formatting: {} }] });

  model.sections = [{ blocks }];
  return model;
}

async function main() {
  const model = build();
  const { bytes } = await writeDocx(model, { modified: new Date("2026-01-01T00:00:00.000Z") });
  const outDir = path.join(__dirname, "..", "e2e", "fixtures");
  mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, "spec-047-fixture.docx");
  writeFileSync(outPath, bytes);
  console.log(`wrote e2e/fixtures/spec-047-fixture.docx (${bytes.length} bytes)`);
}

await main();
// writeDocx's chunked save path yields through idleScheduler's shared
// MessageChannel, whose started port keeps Node's event loop alive after
// the file is fully written — without an explicit exit this script (and
// the benchmark's execFileSync waiting on it) hangs forever.
process.exit(0);
