import { describe, expect, it } from "vitest";
import { createFieldSpecRun, type Block, type DocumentModel, type Paragraph, type Table } from "./documentModel";
import { compileQuery, findMatches, isQueryError, matchSelection } from "./findReplace";
import { documentText, documentTextCaret } from "./documentText";

function paragraph(...texts: string[]): Paragraph {
  return { type: "paragraph", runs: texts.map((text) => ({ text, formatting: {} })) };
}

function table(rows: string[][]): Table {
  return {
    type: "table",
    columnWidthsCm: rows[0].map(() => 3),
    rows: rows.map((cells) => ({ cells: cells.map((text) => ({ blocks: [paragraph(text)] })) })),
  } as Table;
}

function docOf(blocks: Block[], extra: Partial<DocumentModel> = {}): DocumentModel {
  return {
    id: "doc",
    title: "Harbour report",
    sections: [{ blocks }],
    defaultStyles: {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 },
      fontFamily: "Calibri",
      fontSizePt: 11,
    },
    ...extra,
  };
}

describe("documentText", () => {
  it("walks headers, body, footers, then notes and comments, in document order", () => {
    const model = docOf([paragraph("Opening"), table([["A1", "B1"], ["A2", "B2"]]), paragraph("Closing")], {
      notes: [{ noteId: "n1", kind: "footnote", blocks: [paragraph("A footnote")] }],
      comments: [{ threadId: "t1", resolved: false, messages: [{ messageId: "m1", author: "Ana", text: "Check this", createdAt: "2026-10-01T00:00:00Z" }] }],
    });
    model.sections[0].headers = { default: { blocks: [paragraph("Running head")] } };
    model.sections[0].footers = { default: { blocks: [paragraph("Page footer")] } };
    const extracted = documentText(model);
    expect(extracted.coverage).toBe("full");
    expect(extracted.text.split("\n")).toEqual([
      "Running head",
      "Opening",
      "A1",
      "B1",
      "A2",
      "B2",
      "Closing",
      "Page footer",
      "A footnote",
      "Check this",
    ]);
    expect(extracted.spans.map((span) => span.part)).toEqual([
      "header", "body", "body", "body", "body", "body", "body", "footer", "note", "comment",
    ]);
  });

  it("reads table cells row by row, skipping covered slots, and a shape's text frame", () => {
    const merged = table([["Wide", "covered"], ["Left", "Right"]]);
    merged.rows[0].cells[0].colSpan = 2;
    merged.rows[0].cells[1] = { blocks: [], covered: true } as unknown as Table["rows"][number]["cells"][number];
    const shape = { type: "shape", geometry: "rect", text: { vAlign: "top", blocks: [paragraph("Boxed text")] } } as unknown as Block;
    expect(documentText(docOf([merged, shape])).text).toBe("Wide\nLeft\nRight\nBoxed text");
  });

  it("shows a field's value, as the find bar reads it", () => {
    const dated: Paragraph = {
      type: "paragraph",
      runs: [{ text: "Dated ", formatting: {} }, createFieldSpecRun({ kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" }), { text: " today", formatting: {} }],
    };
    expect(documentText(docOf([dated])).text).toBe("Dated 4 March 2026 today");
  });

  it("reports offsets that line up with what matchSelection needs", () => {
    const dated: Paragraph = {
      type: "paragraph",
      runs: [{ text: "Dated ", formatting: {} }, createFieldSpecRun({ kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" }), { text: " Rotterdam", formatting: {} }],
    };
    const model = docOf([paragraph("Intro"), table([["Rotterdam cell"]]), paragraph("The Rotterdam figures"), dated]);
    const extracted = documentText(model);
    const compiled = compileQuery("Rotterdam", { matchCase: false, wholeWords: false, regex: false });
    if (compiled === null || isQueryError(compiled)) throw new Error("query");
    const anchors = findMatches(model, compiled).map((match) => matchSelection(match).anchor);
    // Every body occurrence of the word, located in the extracted text.
    const offsets = [...extracted.text.matchAll(/Rotterdam/g)].map((match) => match.index ?? 0);
    const carets = offsets.map((offset) => documentTextCaret(model, extracted, offset));
    // The table cell maps to its table's start; the paragraphs exactly —
    // across a field whose shown value is longer than its placeholder.
    expect(carets[0]).toEqual({ section: 0, block: 1, offset: 0 });
    expect(carets.slice(1)).toEqual(anchors);
    expect(documentTextCaret(model, extracted, extracted.text.indexOf("March"))).toEqual({ section: 0, block: 3, offset: 6 });
  });

  it("has no caret for text the find bar does not search", () => {
    const model = docOf([paragraph("Body")], { notes: [{ noteId: "n", kind: "endnote", blocks: [paragraph("Endnote")] }] });
    const extracted = documentText(model);
    expect(documentTextCaret(model, extracted, extracted.text.indexOf("Endnote"))).toBeNull();
    expect(documentTextCaret(model, extracted, 999)).toBeNull();
  });

  it("reads an empty document as empty, and full", () => {
    expect(documentText(docOf([]))).toEqual({ text: "", coverage: "full", spans: [] });
  });
});
