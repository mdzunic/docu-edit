import { describe, it, expect } from "vitest";
import { createDocument, type Block, type DocumentModel, type Image, type Table } from "./documentModel";
import { ceilingCounts, ceilingWarningMessage, checkSupportedCeiling } from "./supportedCeiling";

function paragraph(): Block {
  return { type: "paragraph", runs: [] };
}

function image(): Image {
  return {
    type: "image",
    mediaId: "m1",
    naturalWidthPx: 10,
    naturalHeightPx: 10,
    widthCm: 1,
    heightCm: 1,
    wrap: "inline",
  } as Image;
}

function table(cellParagraphs: number): Table {
  return {
    type: "table",
    columnWidthsCm: [5],
    rows: [
      {
        cells: [{ blocks: Array.from({ length: cellParagraphs }, () => ({ type: "paragraph", runs: [] })) }],
      },
    ],
  } as Table;
}

function docWithBlocks(blocks: Block[]): DocumentModel {
  const doc = createDocument();
  doc.sections = [{ blocks }];
  return doc;
}

describe("ceilingCounts", () => {
  it("counts a plain paragraph document as one block, no images/tables/notes", () => {
    const doc = docWithBlocks([paragraph(), paragraph()]);
    expect(ceilingCounts(doc)).toEqual({ blocks: 2, images: 0, tables: 0, notes: 0 });
  });

  it("counts an image block toward both blocks and images", () => {
    const doc = docWithBlocks([image()]);
    expect(ceilingCounts(doc)).toEqual({ blocks: 1, images: 1, tables: 0, notes: 0 });
  });

  it("counts a table as one block plus one block per cell paragraph, recursing into cells", () => {
    const doc = docWithBlocks([table(3)]);
    const counts = ceilingCounts(doc);
    expect(counts.tables).toBe(1);
    // 1 for the table itself + 3 cell paragraphs.
    expect(counts.blocks).toBe(4);
  });

  it("counts footnotes and endnotes together as notes", () => {
    const doc = docWithBlocks([paragraph()]);
    doc.notes = [
      { noteId: "f1", kind: "footnote", blocks: [paragraph()] },
      { noteId: "e1", kind: "endnote", blocks: [paragraph()] },
    ];
    expect(ceilingCounts(doc).notes).toBe(2);
  });
});

describe("checkSupportedCeiling", () => {
  it("reports not exceeded for a document within every limit", () => {
    const doc = docWithBlocks([paragraph(), image(), table(2)]);
    expect(checkSupportedCeiling(doc)).toEqual({ exceeded: false });
  });

  it("reports the images limit exceeded with the value reached", () => {
    const doc = docWithBlocks(Array.from({ length: 1001 }, () => image()));
    const result = checkSupportedCeiling(doc);
    expect(result).toEqual({ exceeded: true, name: "images", limit: 1000, value: 1001 });
  });

  it("reports the tables limit exceeded", () => {
    const doc = docWithBlocks(Array.from({ length: 501 }, () => table(0)));
    const result = checkSupportedCeiling(doc);
    expect(result).toEqual({ exceeded: true, name: "tables", limit: 500, value: 501 });
  });

  it("reports the blocks limit exceeded before a more specific limit when both are over", () => {
    // 26,000 plain paragraphs: over the 25,000 blocks ceiling, nothing else over.
    const doc = docWithBlocks(Array.from({ length: 26_000 }, () => paragraph()));
    const result = checkSupportedCeiling(doc);
    expect(result.name).toBe("blocks");
  });

  it("reports footnotes and endnotes exceeded", () => {
    const doc = docWithBlocks([paragraph()]);
    doc.notes = Array.from({ length: 1001 }, (_, i) => ({
      noteId: `n${i}`,
      kind: "footnote" as const,
      blocks: [paragraph()],
    }));
    const result = checkSupportedCeiling(doc);
    expect(result).toEqual({ exceeded: true, name: "footnotes and endnotes", limit: 1000, value: 1001 });
  });
});

describe("ceilingWarningMessage", () => {
  it("names the limit and value reached, matching the spec's example shape", () => {
    const message = ceilingWarningMessage({ exceeded: true, name: "images", limit: 1000, value: 1240 });
    expect(message).toBe("This document is too large to edit (1,240 images; the limit is 1,000). It's open for reading only.");
  });

  it("is empty for a document within every limit", () => {
    expect(ceilingWarningMessage({ exceeded: false })).toBe("");
  });
});
