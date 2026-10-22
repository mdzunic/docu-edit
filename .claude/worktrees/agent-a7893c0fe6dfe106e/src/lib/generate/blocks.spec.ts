import { describe, expect, it } from "vitest";
import { createDocument, isToc, type DocumentModel } from "../documentModel";
import {
  BODY_BLOCK_TYPES,
  buildBlockFromSpec,
  buildHeaderFooterContent,
  buildHeadingParagraph,
  buildImageBlock,
  buildListParagraphs,
  buildPageBreak,
  buildParagraph,
  buildTable,
  buildTocBlock,
  finalizeTocEntries,
  type BuildContext,
} from "./blocks";
import { DocumentGenerationError } from "./errors";

// 1x1 red PNG.
const PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

function ctx(doc: DocumentModel = createDocument()): BuildContext {
  return { doc };
}

describe("buildHeadingParagraph", () => {
  it("sets headingLevel and bakes the built-in heading style's formatting", () => {
    const paragraph = buildHeadingParagraph(ctx(), "Title", 1, {}, "body[0]");
    expect(paragraph.headingLevel).toBe(1);
    expect(paragraph.runs[0].formatting.bold).toBe(true);
    expect(paragraph.runs[0].formatting.fontSizePt).toBe(20);
    expect(paragraph.formatting?.spaceBeforePt).toBe(20);
  });

  it("an explicit run override wins over the heading style's own formatting", () => {
    const paragraph = buildHeadingParagraph(
      ctx(),
      [{ text: "Title", bold: false, fontSizePt: 30 }],
      1,
      {},
      "body[0]",
    );
    expect(paragraph.runs[0].formatting.bold).toBe(false);
    expect(paragraph.runs[0].formatting.fontSizePt).toBe(30);
  });

  it("rejects a level outside 1..6", () => {
    expect(() => buildHeadingParagraph(ctx(), "Title", 0, {}, "body[0]")).toThrow(DocumentGenerationError);
    expect(() => buildHeadingParagraph(ctx(), "Title", 7, {}, "body[0]")).toThrow(DocumentGenerationError);
  });
});

describe("buildParagraph", () => {
  it("bakes the document's default font family/size onto every run", () => {
    const paragraph = buildParagraph(ctx(), "Hello", {}, "body[0]");
    expect(paragraph.runs[0].formatting.fontFamily).toBe("Calibri");
    expect(paragraph.runs[0].formatting.fontSizePt).toBe(11);
  });

  it("applies alignment/spacing options", () => {
    const paragraph = buildParagraph(
      ctx(),
      "Hello",
      { alignment: "center", spaceBeforePt: 10, spaceAfterPt: 5 },
      "body[0]",
    );
    expect(paragraph.formatting).toEqual({ alignment: "center", spaceBeforePt: 10, spaceAfterPt: 5 });
  });

  it("an empty run array produces an empty paragraph", () => {
    const paragraph = buildParagraph(ctx(), [], {}, "body[0]");
    expect(paragraph.runs).toEqual([]);
  });
});

describe("buildListParagraphs", () => {
  it("builds one paragraph per item, all at level 1", () => {
    const paragraphs = buildListParagraphs(ctx(), ["a", "b", "c"], { style: "bullet" }, "body[0]");
    expect(paragraphs).toHaveLength(3);
    for (const p of paragraphs) {
      expect(p.list).toEqual({ type: "bullet", level: 1 });
    }
  });

  it("defaults to bullet and carries startAt on every item (matching a reopened list's own shape)", () => {
    const paragraphs = buildListParagraphs(ctx(), ["a", "b"], { style: "number", startAt: 5 }, "body[0]");
    expect(paragraphs[0].list?.type).toBe("number");
    expect(paragraphs[0].list?.startAt).toBe(5);
    expect(paragraphs[1].list?.startAt).toBe(5);
  });

  it("rejects an empty items array", () => {
    expect(() => buildListParagraphs(ctx(), [], {}, "body[0]")).toThrow(DocumentGenerationError);
  });
});

describe("buildTable", () => {
  it("builds a rectangular grid with equal column widths by default", () => {
    const table = buildTable(
      ctx(),
      { rows: [["A", "B"], ["1", "2"]] },
      "body[0]",
    );
    expect(table.rows).toHaveLength(2);
    expect(table.columnWidthsCm).toHaveLength(2);
    expect(table.widthMode).toBe("fixed");
    expect(table.rows[0].cells[0].blocks[0].runs[0].text).toBe("A");
  });

  it("marks the first row as a header when header: true", () => {
    const table = buildTable(ctx(), { rows: [["A"], ["B"]], header: true }, "body[0]");
    expect(table.rows[0].isHeader).toBe(true);
    expect(table.rows[1].isHeader).toBeUndefined();
  });

  it("rejects rows with an inconsistent cell count", () => {
    expect(() => buildTable(ctx(), { rows: [["A", "B"], ["C"]] }, "body[0]")).toThrow(
      /same number of cells/,
    );
  });

  it("rejects an unknown property (AC-9)", () => {
    expect(() => buildTable(ctx(), { rows: [["A"]], bogus: 1 }, "body[0]")).toThrow(/unknown property/);
  });

  it("accepts explicit column widths", () => {
    const table = buildTable(
      ctx(),
      { rows: [["A", "B"]], columnWidths: [{ cm: 5 }, { cm: 3 }] },
      "body[0]",
    );
    expect(table.columnWidthsCm[0]).toBeCloseTo(5, 6);
    expect(table.columnWidthsCm[1]).toBeCloseTo(3, 6);
  });

  it("rejects a columnWidths array of the wrong length", () => {
    expect(() =>
      buildTable(ctx(), { rows: [["A", "B"]], columnWidths: [{ cm: 5 }] }, "body[0]"),
    ).toThrow(/columnWidths/);
  });
});

describe("buildImageBlock", () => {
  it("decodes natural dimensions and defaults display size to fit the text area", () => {
    const image = buildImageBlock(ctx(), { data: PNG_BASE64, contentType: "image/png" }, "body[0]");
    expect(image.naturalWidthPx).toBe(1);
    expect(image.naturalHeightPx).toBe(1);
    expect(image.wrap).toBe("inline");
  });

  it("adds a MediaPart to the context's document", () => {
    const doc = createDocument();
    buildImageBlock({ doc }, { data: PNG_BASE64, contentType: "image/png" }, "body[0]");
    expect(doc.media).toHaveLength(1);
    expect(doc.media?.[0].contentType).toBe("image/png");
  });

  it("accepts an explicit width/height", () => {
    const image = buildImageBlock(
      ctx(),
      { data: PNG_BASE64, contentType: "image/png", width: { cm: 4 }, height: { cm: 3 } },
      "body[0]",
    );
    expect(image.widthCm).toBeCloseTo(4, 6);
    expect(image.heightCm).toBeCloseTo(3, 6);
  });

  it("rejects an unsupported content type", () => {
    expect(() =>
      buildImageBlock(ctx(), { data: PNG_BASE64, contentType: "image/svg+xml" }, "body[0]"),
    ).toThrow(DocumentGenerationError);
  });

  it("rejects data that isn't valid base64", () => {
    expect(() =>
      buildImageBlock(ctx(), { data: "not-a-real-image", contentType: "image/png" }, "body[0]"),
    ).toThrow(/not valid base64/);
  });

  it("rejects valid base64 that isn't a valid image for the declared content type", () => {
    expect(() =>
      buildImageBlock(ctx(), { data: btoa("not a png"), contentType: "image/png" }, "body[0]"),
    ).toThrow(/could not decode/);
  });

  it("accepts raw bytes (Uint8Array) as well as base64", () => {
    const bytes = Uint8Array.from(atob(PNG_BASE64), (c) => c.charCodeAt(0));
    const image = buildImageBlock(ctx(), { data: bytes, contentType: "image/png" }, "body[0]");
    expect(image.naturalWidthPx).toBe(1);
  });
});

describe("buildPageBreak", () => {
  it("builds a page break block", () => {
    expect(buildPageBreak()).toEqual({ type: "break", kind: "page" });
  });
});

describe("buildTocBlock", () => {
  it("defaults to levels 1-3, page numbers shown, not hyperlinks-only", () => {
    const toc = buildTocBlock(undefined, "body[0]");
    expect(toc.options).toEqual({ levels: [1, 2, 3], showPageNumbers: true, hyperlinksOnly: false });
    expect(toc.entries).toEqual([]);
  });

  it("rejects an unknown property", () => {
    expect(() => buildTocBlock({ bogus: true }, "body[0]")).toThrow(/unknown property/);
  });
});

describe("finalizeTocEntries", () => {
  it("resolves a TOC block's entries from the document's current headings", () => {
    const doc = createDocument();
    doc.sections[0].blocks = [
      buildHeadingParagraph({ doc }, "Chapter One", 1, {}, "body[0]"),
      buildTocBlock({ levels: [1] }, "body[1]"),
    ];
    finalizeTocEntries(doc);
    const toc = doc.sections[0].blocks[1];
    expect(isToc(toc)).toBe(true);
    if (isToc(toc)) {
      expect(toc.entries).toEqual([
        { text: "Chapter One", level: 1, anchor: expect.any(String), pageNumber: null },
      ]);
    }
  });
});

describe("buildBlockFromSpec (JSON dispatch, AC-4)", () => {
  it("dispatches every body block type", () => {
    const context = ctx();
    expect(buildBlockFromSpec(context, { type: "heading", text: "H", level: 1 }, BODY_BLOCK_TYPES, "body[0]")).toHaveLength(1);
    expect(buildBlockFromSpec(context, { type: "paragraph", text: "p" }, BODY_BLOCK_TYPES, "body[0]")).toHaveLength(1);
    expect(buildBlockFromSpec(context, { type: "list", items: ["a"] }, BODY_BLOCK_TYPES, "body[0]")).toHaveLength(1);
    expect(
      buildBlockFromSpec(context, { type: "table", rows: [["A"]] }, BODY_BLOCK_TYPES, "body[0]"),
    ).toHaveLength(1);
    expect(
      buildBlockFromSpec(
        context,
        { type: "image", data: PNG_BASE64, contentType: "image/png" },
        BODY_BLOCK_TYPES,
        "body[0]",
      ),
    ).toHaveLength(1);
    expect(buildBlockFromSpec(context, { type: "pageBreak" }, BODY_BLOCK_TYPES, "body[0]")).toHaveLength(1);
    expect(buildBlockFromSpec(context, { type: "toc" }, BODY_BLOCK_TYPES, "body[0]")).toHaveLength(1);
  });

  it("rejects a block type outside the allowed set", () => {
    expect(() =>
      buildBlockFromSpec(ctx(), { type: "toc" }, ["heading", "paragraph"], "body[0]"),
    ).toThrow(/expected one of/);
  });

  it("rejects an unknown block type", () => {
    expect(() => buildBlockFromSpec(ctx(), { type: "bogus" }, BODY_BLOCK_TYPES, "body[0]")).toThrow(
      DocumentGenerationError,
    );
  });

  it("rejects a missing type", () => {
    expect(() => buildBlockFromSpec(ctx(), { text: "hi" }, BODY_BLOCK_TYPES, "body[0]")).toThrow(
      DocumentGenerationError,
    );
  });

  it("names the offending path (AC-7)", () => {
    expect(() =>
      buildBlockFromSpec(ctx(), { type: "heading", text: "H" }, BODY_BLOCK_TYPES, "body[3]"),
    ).toThrow(/body\[3\]\.level/);
  });
});

describe("buildHeaderFooterContent", () => {
  it("accepts the plain-string shorthand", () => {
    const content = buildHeaderFooterContent(ctx(), "Company Inc.", "header");
    expect(content.blocks).toHaveLength(1);
    expect(content.blocks[0].type).toBe("paragraph");
  });

  it("accepts a block-spec array with multiple constructs", () => {
    const content = buildHeaderFooterContent(
      ctx(),
      [
        { type: "paragraph", text: "line one" },
        { type: "image", data: PNG_BASE64, contentType: "image/png" },
      ],
      "header",
    );
    expect(content.blocks.map((b) => b.type)).toEqual(["paragraph", "image"]);
  });

  it("rejects a pageBreak/toc block (AC document: header/footer scoped types)", () => {
    expect(() => buildHeaderFooterContent(ctx(), [{ type: "pageBreak" }], "header")).toThrow(
      /expected one of/,
    );
    expect(() => buildHeaderFooterContent(ctx(), [{ type: "toc" }], "header")).toThrow(/expected one of/);
  });
});
