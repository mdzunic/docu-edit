// SPEC-045 AC-11..18: fixtures covering every supported block type, and the
// generate -> write -> re-open round trip asserted via `equalStructural`
// for both formats.
import { describe, expect, it } from "vitest";
import { DocumentBuilder } from "./builder";
import { readDocx, readOdt } from "./read";
import { validateDocx } from "../docx/validate";
import { equalStructural } from "../equality";
import { isImage, isTable, type Block, type DocumentModel } from "../documentModel";

// 1x1 red PNG.
const PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

/** Normalises two known representation-only asymmetries between a
 * freshly-built model and the SAME content re-opened from a saved file,
 * neither of which is content this API's own output ever carries:
 *
 * - `mediaId` is regenerated from the saved package's own part path on
 *   read (documentModel.ts's `MediaPart` doc comment says as much — it is
 *   deliberately not treated like an `id` field structural equality
 *   already ignores, because a WRONG reference must still fail this
 *   check; only the VALUE differs here, remapped to the same order-based
 *   placeholder on both sides so the reference relationship itself is
 *   still verified).
 * - a table's `docxTblPr`/`docxTrPr`/`docxTcPr`/`odtStyleName`/`odtName`/
 *   `odtColumnStyleNames` are SPEC-008 preservation slots for a source
 *   file's unmodelled table styling — both readers always populate them
 *   on open, even for a table with nothing unmodelled to preserve. A
 *   paragraph's `styleId` picks up the same kind of reader-only artifact
 *   from the .odt writer's own automatic-style-per-paragraph scheme (e.g.
 *   the style that carries a page break's `fo:break-before`) — this API
 *   never sets `styleId` itself (see builder.ts's doc comment), so any
 *   value found here on the reopened side is one of the writer's own
 *   bookkeeping choices, not model content. */
function normalizeForRoundTrip(doc: DocumentModel): DocumentModel {
  const clone = structuredClone(doc);
  const mediaIdMap = new Map<string, string>();
  (clone.media ?? []).forEach((part, i) => {
    mediaIdMap.set(part.mediaId, `media-${i}`);
    part.mediaId = `media-${i}`;
  });
  const strip = (block: Block): void => {
    if (block.type === "paragraph") {
      delete block.styleId;
    }
    if (isImage(block)) {
      block.mediaId = mediaIdMap.get(block.mediaId) ?? block.mediaId;
    }
    if (isTable(block)) {
      delete block.docxTblPr;
      delete block.odtStyleName;
      delete block.odtName;
      delete block.odtColumnStyleNames;
      for (const row of block.rows) {
        delete row.docxTrPr;
        delete row.odtStyleName;
        for (const cell of row.cells) {
          delete cell.docxTcPr;
          delete cell.odtStyleName;
          cell.blocks.forEach(strip);
        }
      }
    }
  };
  for (const section of clone.sections) {
    section.blocks.forEach(strip);
    for (const bag of [section.headers, section.footers]) {
      for (const content of Object.values(bag ?? {})) {
        content?.blocks.forEach(strip);
      }
    }
  }
  return clone;
}

/** One document exercising every block type AC-13 names (heading,
 * paragraph with formatted runs, list, table, image, page break, header,
 * footer, TOC). The page break sits mid-document, followed by a real
 * paragraph — a page break with nothing after it is a documented .odt
 * writer edge case (a synthetic carrier paragraph, see docs/generate-api.md)
 * this fixture does not need to exercise. The TOC uses `hyperlinksOnly`
 * with no shown page numbers: this API never has real pagination to cache,
 * and both readers derive `showPageNumbers`/`hyperlinksOnly` from whether
 * any entry actually carries a cached page number on reopen — an entry
 * this API writes never does, so a fixture claiming `showPageNumbers: true`
 * could never itself round-trip that claim. */
function buildFixture(): DocumentBuilder {
  return new DocumentBuilder({ title: "Fixture" })
    .addHeading("Chapter One", 1)
    .addHeading("Section A", 2)
    .addParagraph([{ text: "bold ", bold: true }, { text: "italic", italic: true }, " plain"])
    .addList(["First", "Second", "Third"], { style: "bullet" })
    // A plain paragraph between two addList calls of a different style is
    // required, not cosmetic: lib/lists.ts's planListLevels defines each
    // outline level from the FIRST item at that level in a contiguous run
    // of list paragraphs (the shared multilevel-list numbering plan both
    // writers use) — two directly-adjacent lists of different style at the
    // same level would otherwise collapse into one, taking the first
    // list's style (see docs/generate-api.md).
    .addParagraph("Numbered from three:")
    .addList(["One", "Two"], { style: "number", startAt: 3 })
    .addTable({
      rows: [
        ["Item", "Qty"],
        ["Widget", "3"],
      ],
      header: true,
    })
    .addImage({ data: PNG_BASE64, contentType: "image/png", altText: "A red dot" })
    .addTableOfContents({ levels: [1, 2], hyperlinksOnly: true, showPageNumbers: false })
    .addPageBreak()
    .addParagraph("After the break")
    .setHeader("Fixture header")
    .setFooter("Fixture footer");
}

describe("generate -> write -> re-open (AC-13, AC-14, AC-15, AC-18)", () => {
  it(".docx: equalStructural holds for the body, headers, footers and media", async () => {
    const builder = buildFixture();
    const bytes = await builder.toBytes("docx");
    const reopened = await readDocx(bytes);
    const generated = normalizeForRoundTrip(builder.toDoc());
    const reopenedNormalized = normalizeForRoundTrip(reopened);
    expect(equalStructural(generated.sections, reopenedNormalized.sections)).toBe(true);
    expect(equalStructural(generated.media, reopenedNormalized.media)).toBe(true);
  });

  it(".odt: equalStructural holds for the body, headers, footers and media", async () => {
    const builder = buildFixture();
    const bytes = await builder.toBytes("odt");
    const reopened = await readOdt(bytes);
    const generated = normalizeForRoundTrip(builder.toDoc());
    const reopenedNormalized = normalizeForRoundTrip(reopened);
    expect(equalStructural(generated.sections, reopenedNormalized.sections)).toBe(true);
    expect(equalStructural(generated.media, reopenedNormalized.media)).toBe(true);
  });

  // validateDocx's schema allowlist (docx/validate.ts) predates images and
  // headers/footers (it is never exercised against w:drawing or
  // w:headerReference/w:footerReference by that spec's own test suites
  // either — see docx/images.spec.ts, docx/headers.spec.ts), so this
  // checks the constructs it does cover; the image/header/footer fixture
  // above is verified by the round-trip assertion instead, the same
  // precedent those specs' own tests set.
  it("the generated .docx passes the writer's own OOXML schema validator (AC-12)", async () => {
    const bytes = await new DocumentBuilder({ title: "Fixture" })
      .addHeading("Chapter One", 1)
      .addParagraph([{ text: "bold ", bold: true }, "plain"])
      .addList(["First", "Second"], { style: "bullet" })
      .addTable({ rows: [["A", "B"]], header: true })
      .addTableOfContents({ hyperlinksOnly: true, showPageNumbers: false })
      .addPageBreak()
      .addParagraph("after")
      .toBytes("docx");
    expect(await validateDocx(bytes)).toEqual([]);
  });
});

describe("readDocx / readOdt (AC-17)", () => {
  it("accepts a Buffer with no disk round trip", async () => {
    const builder = buildFixture();
    const docxBuffer = await builder.toBuffer("docx");
    expect(Buffer.isBuffer(docxBuffer)).toBe(true);
    const docxModel = await readDocx(docxBuffer);
    expect(docxModel.sections[0].blocks.length).toBeGreaterThan(0);

    const odtBuffer = await builder.toBuffer("odt");
    const odtModel = await readOdt(odtBuffer);
    expect(odtModel.sections[0].blocks.length).toBeGreaterThan(0);
  });
});

describe("empty document (AC-10, AC-11)", () => {
  it("a document with zero body blocks is accepted", () => {
    const doc = new DocumentBuilder().toDoc();
    expect(doc.sections[0].blocks).toEqual([]);
  });

  it("produces a valid, openable .docx with no repair prompt", async () => {
    const bytes = await new DocumentBuilder().toBytes("docx");
    expect(await validateDocx(bytes)).toEqual([]);
    const reopened = await readDocx(bytes);
    expect(reopened.sections[0].blocks).toHaveLength(1);
    expect(reopened.sections[0].blocks[0]).toMatchObject({ type: "paragraph" });
  });

  it("produces a valid, openable .odt", async () => {
    const bytes = await new DocumentBuilder().toBytes("odt");
    const reopened = await readOdt(bytes);
    expect(reopened.sections[0].blocks).toHaveLength(1);
    expect(reopened.sections[0].blocks[0]).toMatchObject({ type: "paragraph" });
  });
});
