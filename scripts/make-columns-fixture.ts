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

const { bytes } = await writeDocx(model);
writeFileSync("e2e/fixtures/two-column-section.docx", bytes);
console.log(`wrote e2e/fixtures/two-column-section.docx (${bytes.length} bytes)`);
