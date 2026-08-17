#!/usr/bin/env -S npx vite-node
// One-off generator for the SPEC-031 footnote fixture (run manually:
// `npx vite-node scripts/build-spec031-fixtures.mjs`) — produced through
// this project's OWN writer, like the SPEC-029 fixtures: foreign-producer
// XML shapes are covered by hand-authored unit tests
// (src/lib/docx/notes.spec.ts); what this fixture exercises is AC-25's
// end-to-end open -> paginate -> place-notes pipeline, which needs real
// multi-page package bytes with 20+ footnotes.
//
// Four pages forced by manual page breaks (deterministic regardless of font
// metrics), six footnotes per page — 24 in all, each referenced from its
// own body paragraph so every note's page is unambiguous.
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { createDocument, createNoteRun } from "../src/lib/documentModel.ts";
import { writeDocx } from "../src/lib/docx/writer.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(e2eFixturesDir, { recursive: true });

const PAGES = 4;
const NOTES_PER_PAGE = 6;

// Pinned so a regenerated fixture's dcterms:modified stays put instead of
// becoming whichever day someone re-ran this script — otherwise regenerating
// churns the committed bytes for a reason unrelated to the change being made
// (SPEC-008's round-trip determinism rule).
const FIXED_MODIFIED = new Date("2026-01-01T00:00:00.000Z");

function model() {
  const doc = createDocument("spec031-footnotes-fixture");
  doc.title = "Twenty-four footnotes";
  const blocks = [];
  const notes = [];
  let noteNumber = 0;
  for (let pageIndex = 0; pageIndex < PAGES; pageIndex += 1) {
    if (pageIndex > 0) {
      blocks.push({ type: "break", kind: "page" });
    }
    for (let slot = 0; slot < NOTES_PER_PAGE; slot += 1) {
      noteNumber += 1;
      const noteId = `fixture-note-${noteNumber}`;
      blocks.push({
        type: "paragraph",
        runs: [
          { text: `Page ${pageIndex + 1} paragraph ${slot + 1} with a reference`, formatting: {} },
          createNoteRun(noteId, "footnote"),
        ],
      });
      notes.push({
        noteId,
        kind: "footnote",
        blocks: [
          { type: "paragraph", runs: [{ text: `Note ${noteNumber} content`, formatting: {} }] },
        ],
      });
    }
  }
  doc.sections = [{ blocks }];
  doc.notes = notes;
  return doc;
}

async function main() {
  const { bytes } = await writeDocx(model(), { modified: FIXED_MODIFIED });
  const target = path.join(e2eFixturesDir, "spec031-footnotes.docx");
  writeFileSync(target, bytes);
  console.log(`wrote ${target} (${bytes.length} bytes)`);
}

await main();
// writeDocx's chunked save path yields through idleScheduler's shared
// MessageChannel, whose started port keeps Node's event loop alive after the
// file is fully written — without an explicit exit this script hangs forever
// (same reason build-spec047-fixture.mjs ends this way).
process.exit(0);
