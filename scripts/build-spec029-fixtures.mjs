#!/usr/bin/env -S npx vite-node
// One-off generator for the SPEC-029 header/footer fixtures (run manually:
// `npx vite-node scripts/build-spec029-fixtures.mjs`). Unlike SPEC-024's
// hand-built-byte-for-byte fixtures, these are produced through this
// project's OWN writer (writeDocx/writeOdt) — the reader's compatibility
// with a genuinely foreign producer's header/footer XML shapes is already
// covered by a hand-authored-XML unit test (src/lib/docx/headers.spec.ts's
// "hand-authored Word-shaped fldSimple" case); what these fixtures exercise
// is the end-to-end open -> paginate -> render pipeline for AC-27/AC-28
// ("opening a fixture with different first-page / odd-even headers renders
// them on the right pages"), which needs real multi-page package bytes, not
// another unit-level parse assertion.
//
// Each fixture forces exactly 3 pages via manual page breaks (deterministic
// regardless of font metrics) so a page-1/page-2/page-3 comparison in the
// e2e spec is meaningful.
import { writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { createDocument } from "../src/lib/documentModel.ts";
import { writeDocx } from "../src/lib/docx/writer.ts";
import { writeOdt } from "../src/lib/odf/writer.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const e2eFixturesDir = path.join(__dirname, "..", "e2e", "fixtures");
mkdirSync(e2eFixturesDir, { recursive: true });

// Pinned so a regenerated fixture's dcterms:modified / dc:date stays put
// instead of becoming whichever day someone re-ran this script — otherwise
// regenerating churns the committed bytes for a reason unrelated to the
// change being made (SPEC-008's round-trip determinism rule).
const FIXED_MODIFIED = new Date("2026-01-01T00:00:00.000Z");

function paragraph(text) {
  return { type: "paragraph", runs: text ? [{ text, formatting: {} }] : [] };
}

function pageBreak() {
  return { type: "break", kind: "page" };
}

/** Three pages of body text separated by manual page breaks. */
function threePageBody(label) {
  return [
    paragraph(`${label} — page 1 body`),
    pageBreak(),
    paragraph(`${label} — page 2 body`),
    pageBreak(),
    paragraph(`${label} — page 3 body`),
  ];
}

function firstPageModel(id) {
  const model = createDocument(id);
  model.title = "Different first page";
  model.sections = [
    {
      blocks: threePageBody("First-page fixture"),
      titlePage: true,
      headers: {
        first: { blocks: [paragraph("COVER PAGE")] },
        default: { blocks: [paragraph("Regular header")] },
      },
    },
  ];
  return model;
}

function oddEvenModel(id) {
  const model = createDocument(id);
  model.title = "Odd and even headers";
  model.sections = [
    {
      blocks: threePageBody("Odd/even fixture"),
      evenAndOddHeaders: true,
      headers: {
        default: { blocks: [paragraph("ODD HEADER")] },
        even: { blocks: [paragraph("EVEN HEADER")] },
      },
    },
  ];
  return model;
}

async function main() {
  // "Different first page" (AC-27) is .docx-only in this writer — ODF has
  // no direct equivalent of w:titlePg (a real producer represents it with a
  // second master page linked via style:next-style-name, out of scope for
  // this pass, see odf/writer.ts's header comments) — so only a .docx
  // fixture demonstrates it.
  const firstPageDocx = await writeDocx(firstPageModel("spec029-first-page-docx"), { modified: FIXED_MODIFIED });
  writeFileSync(path.join(e2eFixturesDir, "spec029-first-page.docx"), firstPageDocx.bytes);

  const oddEvenDocx = await writeDocx(oddEvenModel("spec029-odd-even-docx"), { modified: FIXED_MODIFIED });
  writeFileSync(path.join(e2eFixturesDir, "spec029-odd-even.docx"), oddEvenDocx.bytes);

  const oddEvenOdt = await writeOdt(oddEvenModel("spec029-odd-even-odt"), { modified: FIXED_MODIFIED });
  writeFileSync(path.join(e2eFixturesDir, "spec029-odd-even.odt"), oddEvenOdt.bytes);

  console.log("Wrote SPEC-029 fixtures to", e2eFixturesDir);
}

await main();
// writeDocx/writeOdt's chunked save path yields through idleScheduler's
// shared MessageChannel, whose started port keeps Node's event loop alive
// after the files are fully written — without an explicit exit this script
// hangs forever (same reason build-spec047-fixture.mjs ends this way).
process.exit(0);
