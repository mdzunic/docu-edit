import { describe, it, expect } from "vitest";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { blockText, type CaretPosition } from "./textEditing";
import {
  blockCount,
  blockOrdinal,
  caretSelection,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentEnd,
  documentStart,
  insertLineBreak,
  insertText,
  isCollapsed,
  orderedRange,
  selectAll,
  splitBlock,
  type ModelSelection,
} from "./commands";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function doc(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("commands-test");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

/** Every command must leave its input untouched (AC-25/AC-26): snapshot the
 * input, run the command, and require the input to still match. */
function snapshot(model: DocumentModel): string {
  return JSON.stringify(model);
}

describe("insertText", () => {
  it("inserts at a collapsed selection and returns the caret after the text (AC-3)", () => {
    const model = doc(paragraph(["Hello world"]));
    const { model: next, selection } = insertText(model, caretSelection(caret(0, 5)), " brave");
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello brave world");
    expect(selection.focus).toEqual(caret(0, 11));
    expect(isCollapsed(selection)).toBe(true);
  });

  it("returns a new model state and never mutates the input (AC-24, AC-25)", () => {
    const model = doc(paragraph(["abc"]));
    const before = snapshot(model);
    const { model: next } = insertText(model, caretSelection(caret(0, 1)), "X");
    expect(next).not.toBe(model);
    expect(snapshot(model)).toBe(before);
    expect(blockText(model.sections[0].blocks[0])).toBe("abc");
    expect(blockText(next.sections[0].blocks[0])).toBe("aXbc");
  });

  it("shares untouched blocks with the input model by reference", () => {
    const untouched = paragraph(["stays"]);
    const model = doc(paragraph(["edited"]), untouched);
    const { model: next } = insertText(model, caretSelection(caret(0, 0)), "x");
    expect(next.sections[0].blocks[1]).toBe(untouched);
    expect(next.sections[0].blocks[0]).not.toBe(model.sections[0].blocks[0]);
  });

  it("replaces a non-collapsed selection with the typed text (AC-23)", () => {
    const model = doc(paragraph(["Hello world"]));
    const { model: next, selection } = insertText(model, range(caret(0, 6), caret(0, 11)), "there");
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello there");
    expect(selection.focus).toEqual(caret(0, 11));
  });

  it("replaces a selection given in reverse (focus before anchor)", () => {
    const model = doc(paragraph(["Hello world"]));
    const { model: next } = insertText(model, range(caret(0, 11), caret(0, 6)), "there");
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello there");
  });

  it("replaces a selection spanning paragraphs, merging them (AC-22, AC-23)", () => {
    const model = doc(paragraph(["First line"]), paragraph(["Second line"]));
    const { model: next, selection } = insertText(model, range(caret(0, 6), caret(1, 6)), "-");
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("First - line");
    expect(selection.focus).toEqual(caret(0, 7));
  });

  it("keeps the formatting of the run the text lands in", () => {
    const model = doc(paragraph(["plain ", {}], ["bold", { bold: true }]));
    const { model: next } = insertText(model, caretSelection(caret(0, 8)), "er");
    const runs = next.sections[0].blocks[0].runs;
    expect(runs).toEqual([
      { text: "plain ", formatting: {} },
      { text: "boerld", formatting: { bold: true } },
    ]);
  });
});

describe("insertLineBreak", () => {
  it("inserts a literal newline without creating a paragraph (AC-10)", () => {
    const model = doc(paragraph(["one two"]));
    const { model: next, selection } = insertLineBreak(model, caretSelection(caret(0, 3)));
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("one\n two");
    expect(selection.focus).toEqual(caret(0, 4));
  });
});

describe("splitBlock", () => {
  it("splits the paragraph at the caret (AC-8)", () => {
    const model = doc(paragraph(["Hello world"]));
    const { model: next, selection } = splitBlock(model, caretSelection(caret(0, 5)));
    expect(next.sections[0].blocks).toHaveLength(2);
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello");
    expect(blockText(next.sections[0].blocks[1])).toBe(" world");
    expect(selection.focus).toEqual(caret(1, 0));
    expect(blockText(model.sections[0].blocks[0])).toBe("Hello world");
  });

  it("carries paragraph formatting to the new paragraph (AC-9)", () => {
    const model = doc(paragraph(["Centered text"]));
    model.sections[0].blocks[0] = {
      ...model.sections[0].blocks[0],
      formatting: { alignment: "center", indentLeftCm: 1 },
      styleId: "Fancy",
      headingLevel: 2,
    };
    const { model: next } = splitBlock(model, caretSelection(caret(0, 8)));
    const [first, second] = next.sections[0].blocks;
    expect(first.formatting).toEqual({ alignment: "center", indentLeftCm: 1 });
    expect(second.formatting).toEqual({ alignment: "center", indentLeftCm: 1 });
    expect(second.styleId).toBe("Fancy");
    expect(second.headingLevel).toBe(2);
  });

  it("deletes a non-collapsed selection before splitting (AC-23)", () => {
    const model = doc(paragraph(["Hello brave world"]));
    const { model: next } = splitBlock(model, range(caret(0, 5), caret(0, 11)));
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello");
    expect(blockText(next.sections[0].blocks[1])).toBe(" world");
  });
});

describe("deleteRange", () => {
  it("deletes within a paragraph and collapses to the start", () => {
    const model = doc(paragraph(["Hello brave world"]));
    const { model: next, selection } = deleteRange(model, range(caret(0, 5), caret(0, 11)));
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello world");
    expect(selection.focus).toEqual(caret(0, 5));
  });

  it("merges paragraphs when the range spans a boundary (AC-7, AC-22)", () => {
    const model = doc(paragraph(["First"]), paragraph(["Second"]), paragraph(["Third"]));
    const before = snapshot(model);
    const { model: next } = deleteRange(model, range(caret(0, 3), caret(2, 2)));
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("Firird");
    expect(snapshot(model)).toBe(before);
  });

  it("is a no-op for a collapsed selection but still returns a new state", () => {
    const model = doc(paragraph(["abc"]));
    const { model: next } = deleteRange(model, caretSelection(caret(0, 1)));
    expect(next).not.toBe(model);
    expect(blockText(next.sections[0].blocks[0])).toBe("abc");
  });
});

describe("deleteBackward", () => {
  it("removes the character before the caret (AC-5)", () => {
    const model = doc(paragraph(["abc"]));
    const { model: next, selection } = deleteBackward(model, caretSelection(caret(0, 2)));
    expect(blockText(next.sections[0].blocks[0])).toBe("ac");
    expect(selection.focus).toEqual(caret(0, 1));
    expect(blockText(model.sections[0].blocks[0])).toBe("abc");
  });

  it("merges with the previous paragraph at a paragraph start (AC-7)", () => {
    const model = doc(paragraph(["First"]), paragraph(["Second"]));
    const { model: next, selection } = deleteBackward(model, caretSelection(caret(1, 0)));
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("FirstSecond");
    expect(selection.focus).toEqual(caret(0, 5));
  });

  it("is a no-op at the document start", () => {
    const model = doc(paragraph(["abc"]));
    const { model: next, selection } = deleteBackward(model, caretSelection(caret(0, 0)));
    expect(blockText(next.sections[0].blocks[0])).toBe("abc");
    expect(selection.focus).toEqual(caret(0, 0));
  });

  it("deletes only the selection when it is non-collapsed", () => {
    const model = doc(paragraph(["Hello world"]));
    const { model: next } = deleteBackward(model, range(caret(0, 5), caret(0, 11)));
    expect(blockText(next.sections[0].blocks[0])).toBe("Hello");
  });
});

describe("deleteForward", () => {
  it("removes the character after the caret (AC-6)", () => {
    const model = doc(paragraph(["abc"]));
    const { model: next, selection } = deleteForward(model, caretSelection(caret(0, 1)));
    expect(blockText(next.sections[0].blocks[0])).toBe("ac");
    expect(selection.focus).toEqual(caret(0, 1));
  });

  it("merges with the next paragraph at a paragraph end (AC-7)", () => {
    const model = doc(paragraph(["First"]), paragraph(["Second"]));
    const { model: next, selection } = deleteForward(model, caretSelection(caret(0, 5)));
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("FirstSecond");
    expect(selection.focus).toEqual(caret(0, 5));
  });

  it("is a no-op at the document end", () => {
    const model = doc(paragraph(["abc"]));
    const { model: next } = deleteForward(model, caretSelection(caret(0, 3)));
    expect(blockText(next.sections[0].blocks[0])).toBe("abc");
  });
});

describe("selection helpers", () => {
  it("documentStart / documentEnd address the model's extremes (AC-13)", () => {
    const model = doc(paragraph(["one"]), paragraph(["three"]));
    expect(documentStart()).toEqual(caret(0, 0));
    expect(documentEnd(model)).toEqual(caret(1, 5));
  });

  it("documentEnd reaches into the last section of a multi-section model", () => {
    const model = doc(paragraph(["one"]));
    model.sections.push({ blocks: [paragraph(["last!"])] });
    expect(documentEnd(model)).toEqual(caret(0, 5, 1));
  });

  it("selectAll spans the whole document (AC-19)", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]));
    expect(selectAll(model)).toEqual({ anchor: caret(0, 0), focus: caret(1, 3) });
  });

  it("orderedRange orders and clamps the ends", () => {
    const model = doc(paragraph(["abc"]), paragraph(["def"]));
    const ordered = orderedRange(model, range(caret(1, 99), caret(0, 1)));
    expect(ordered).toEqual({ start: caret(0, 1), end: caret(1, 3) });
  });

  it("blockCount / blockOrdinal count across sections", () => {
    const model = doc(paragraph(["a"]), paragraph(["b"]));
    model.sections.push({ blocks: [paragraph(["c"])] });
    expect(blockCount(model)).toBe(3);
    expect(blockOrdinal(model, caret(1, 0))).toBe(2);
    expect(blockOrdinal(model, caret(0, 0, 1))).toBe(3);
  });
});
