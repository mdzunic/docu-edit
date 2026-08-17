// One-off generator for the SPEC-030 AC-19 fixture: a .docx whose first
// section is two equal columns of body text and whose second section is a
// single-column heading area. Run with `npx vite-node scripts/make-columns-fixture.ts`.
import { writeFileSync } from "node:fs";
import { writeDocx } from "../src/lib/docx/writer";
import { createDocument, type DocumentModel, type Paragraph } from "../src/lib/documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function heading(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], headingLevel: 1 };
}

const model: DocumentModel = createDocument("two-column-fixture");
model.format = "docx";
model.title = "Two column section";
model.sections = [
  {
    columns: { count: 2, equalWidth: true, gapCm: 1.25, separator: true },
    blocks: Array.from({ length: 8 }, (_, i) => paragraph(`Column paragraph ${i + 1}.`)),
  },
  {
    start: "continuous",
    blocks: [heading("Newsletter Title"), paragraph("Single column footer text.")],
  },
];

// `modified` is pinned so a regenerated fixture's dcterms:modified stays put
// instead of becoming whichever day someone re-ran this script — otherwise
// regenerating churns the committed bytes for a reason unrelated to the
// change being made (SPEC-008's round-trip determinism rule).
const { bytes } = await writeDocx(model, { modified: new Date("2026-01-01T00:00:00.000Z") });
writeFileSync("e2e/fixtures/two-column-section.docx", bytes);
console.log(`wrote e2e/fixtures/two-column-section.docx (${bytes.length} bytes)`);
// writeDocx's chunked save path yields through idleScheduler's shared
// MessageChannel, whose started port keeps Node's event loop alive after the
// file is fully written — without an explicit exit this script hangs forever
// (same reason build-spec047-fixture.mjs ends this way).
process.exit(0);
