// SPEC-028 — manual breaks, section breaks, per-section page setup and the
// paragraph pagination flags, exercised through the command layer.
import { describe, expect, it } from "vitest";
import {
  applyParagraphFormat,
  caretSelection,
  deleteBackward,
  deleteForward,
  deleteRange,
  insertBreak,
  insertSectionBreak,
  paragraphPaginationState,
  setPageSetup,
  type ModelSelection,
} from "./commands";
import {
  createDocument,
  isBreak,
  sectionMargins,
  sectionPageSize,
  type DocumentModel,
  type Paragraph,
} from "./documentModel";
import type { CaretPosition } from "./textEditing";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text ? [{ text, formatting: {} }] : [] };
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function fixture(...texts: string[]): DocumentModel {
  const model = createDocument("breaks-fixture");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

function blockTypes(model: DocumentModel, section = 0): string[] {
  return model.sections[section].blocks.map((block) => block.type);
}

function blockTexts(model: DocumentModel, section = 0): string[] {
  return model.sections[section].blocks.map((block) =>
    block.type === "paragraph" ? block.runs.map((run) => run.text).join("") : `<${block.type}>`,
  );
}

describe("insertBreak", () => {
  it("splits the paragraph around a page break mid-paragraph", () => {
    const { model, selection } = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "page");
    expect(blockTypes(model)).toEqual(["paragraph", "break", "paragraph"]);
    expect(blockTexts(model)).toEqual(["Hello", "<break>", "World"]);
    expect(model.sections[0].blocks[1]).toEqual({ type: "break", kind: "page" });
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("inserts a column break node into the document model (AC-10)", () => {
    const { model, selection } = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "column");
    expect(blockTypes(model)).toEqual(["paragraph", "break", "paragraph"]);
    expect(model.sections[0].blocks[1]).toEqual({ type: "break", kind: "column" });
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("does not leave an empty half when the caret is at a paragraph start", () => {
    const { model, selection } = insertBreak(fixture("One", "Two"), caretSelection(caret(1, 0)), "page");
    expect(blockTexts(model)).toEqual(["One", "<break>", "Two"]);
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("leaves an empty paragraph after a break at the document end", () => {
    const { model, selection } = insertBreak(fixture("End"), caretSelection(caret(0, 3)), "page");
    expect(blockTexts(model)).toEqual(["End", "<break>", ""]);
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("leaves an empty paragraph before a break at the document start", () => {
    // Without it the first page holds nothing but the break: it renders as
    // an empty page no click or arrow key can reach, so its content could
    // never be typed.
    const { model, selection } = insertBreak(fixture("Start"), caretSelection(caret(0, 0)), "page");
    expect(blockTexts(model)).toEqual(["", "<break>", "Start"]);
    expect(selection.focus).toEqual(caret(2, 0));
  });

  it("leaves an empty paragraph before a break at the start of a next-page section", () => {
    const model = fixture("One");
    model.sections.push({ blocks: [paragraph("Two")], start: "nextPage" });
    const result = insertBreak(model, caretSelection({ section: 1, block: 0, offset: 0 }), "page");
    expect(blockTexts(result.model, 1)).toEqual(["", "<break>", "Two"]);
    expect(result.selection.focus).toEqual({ section: 1, block: 2, offset: 0 });
  });

  it("adds no leading paragraph when the section continues the current page", () => {
    const model = fixture("One");
    model.sections.push({ blocks: [paragraph("Two")], start: "continuous" });
    const result = insertBreak(model, caretSelection({ section: 1, block: 0, offset: 0 }), "page");
    expect(blockTexts(result.model, 1)).toEqual(["<break>", "Two"]);
    expect(result.selection.focus).toEqual({ section: 1, block: 1, offset: 0 });
  });

  it("keeps the split paragraph's formatting on both halves", () => {
    const model = fixture("HelloWorld");
    (model.sections[0].blocks[0] as Paragraph).formatting = { alignment: "center" };
    const { model: next } = insertBreak(model, caretSelection(caret(0, 5)), "page");
    expect((next.sections[0].blocks[0] as Paragraph).formatting).toEqual({ alignment: "center" });
    expect((next.sections[0].blocks[2] as Paragraph).formatting).toEqual({ alignment: "center" });
  });

  it("replaces a non-collapsed selection", () => {
    const { model } = insertBreak(fixture("HelloWorld"), range(caret(0, 2), caret(0, 8)), "page");
    expect(blockTexts(model)).toEqual(["He", "<break>", "ld"]);
  });

  it("is inert inside a table cell", () => {
    const model = createDocument("breaks-cell");
    model.sections = [
      {
        blocks: [
          {
            type: "table",
            rows: [{ cells: [{ blocks: [paragraph("cell")] }] }],
            columnWidthsCm: [5],
          },
        ],
      },
    ];
    const inCell: CaretPosition = { section: 0, block: 0, offset: 2, cell: { row: 0, col: 0, block: 0 } };
    const { model: next } = insertBreak(model, caretSelection(inCell), "page");
    expect(blockTypes(next)).toEqual(["table"]);
  });

  it("is undone as one edit: the break and the split disappear together", () => {
    const before = fixture("HelloWorld");
    const { model } = insertBreak(before, caretSelection(caret(0, 5)), "page");
    // Immutability: the input model is untouched, which is what snapshot
    // undo restores.
    expect(blockTexts(before)).toEqual(["HelloWorld"]);
    expect(blockTexts(model)).toEqual(["Hello", "<break>", "World"]);
  });
});

describe("deleting breaks rejoins the flow", () => {
  it("Backspace after a break removes the break and keeps the paragraph", () => {
    const withBreak = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "page").model;
    const { model, selection } = deleteBackward(withBreak, caretSelection(caret(2, 0)));
    expect(blockTexts(model)).toEqual(["Hello", "World"]);
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("a second Backspace then merges the paragraphs back together", () => {
    const withBreak = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "page").model;
    const once = deleteBackward(withBreak, caretSelection(caret(2, 0)));
    const twice = deleteBackward(once.model, once.selection);
    expect(blockTexts(twice.model)).toEqual(["HelloWorld"]);
  });

  it("Delete before a break removes the break", () => {
    const withBreak = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "page").model;
    const { model, selection } = deleteForward(withBreak, caretSelection(caret(0, 5)));
    expect(blockTexts(model)).toEqual(["Hello", "World"]);
    expect(selection.focus).toEqual(caret(0, 5));
  });

  it("a selection spanning a break deletes it with the covered text", () => {
    const withBreak = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "page").model;
    const { model } = deleteRange(withBreak, range(caret(0, 3), caret(2, 2)));
    expect(blockTexts(model)).toEqual(["Helrld"]);
  });
});

describe("insertSectionBreak", () => {
  it("splits the section at the caret and marks how the new one starts", () => {
    const { model, selection } = insertSectionBreak(
      fixture("HelloWorld"),
      caretSelection(caret(0, 5)),
      "nextPage",
    );
    expect(model.sections).toHaveLength(2);
    expect(blockTexts(model, 0)).toEqual(["Hello"]);
    expect(blockTexts(model, 1)).toEqual(["World"]);
    expect(model.sections[1].start).toBe("nextPage");
    expect(selection.focus).toEqual({ section: 1, block: 0, offset: 0 });
  });

  it("supports a continuous start", () => {
    const { model } = insertSectionBreak(fixture("AB"), caretSelection(caret(0, 1)), "continuous");
    expect(model.sections[1].start).toBe("continuous");
  });

  it("the new section inherits the split section's own page setup", () => {
    const model = fixture("HelloWorld");
    model.sections[0].pageSize = { widthCm: 14.8, heightCm: 21 };
    model.sections[0].margins = { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 };
    const { model: next } = insertSectionBreak(model, caretSelection(caret(0, 5)), "nextPage");
    expect(next.sections[1].pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
    expect(next.sections[1].margins).toEqual({ topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 });
  });

  it("keeps a paragraph on both sides when splitting at a section edge", () => {
    const { model } = insertSectionBreak(fixture("Only"), caretSelection(caret(0, 0)), "nextPage");
    expect(blockTexts(model, 0)).toEqual([""]);
    expect(blockTexts(model, 1)).toEqual(["Only"]);
  });

  it("Backspace at the section start deletes the section break (merges back)", () => {
    const split = insertSectionBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "nextPage");
    const { model } = deleteBackward(split.model, split.selection);
    expect(model.sections).toHaveLength(1);
    expect(blockTexts(model)).toEqual(["HelloWorld"]);
  });

  it("Backspace merges sections without eating a trailing break atom", () => {
    const withBreak = insertBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "page").model;
    // Section break right after the page break block: section 1 = [Hello, break], section 2 = [World]
    const split = insertSectionBreak(withBreak, caretSelection(caret(2, 0)), "nextPage");
    expect(blockTypes(split.model, 0)).toEqual(["paragraph", "break"]);
    const merged = deleteBackward(split.model, split.selection);
    expect(merged.model.sections).toHaveLength(1);
    expect(blockTexts(merged.model)).toEqual(["Hello", "<break>", "World"]);
    expect(isBreak(merged.model.sections[0].blocks[1])).toBe(true);
  });

  it("merging keeps the earlier section's own page setup", () => {
    const model = fixture("HelloWorld");
    model.sections[0].pageSize = { widthCm: 14.8, heightCm: 21 };
    const split = insertSectionBreak(model, caretSelection(caret(0, 5)), "nextPage");
    const changed = setPageSetup(
      split.model,
      split.selection,
      {
        pageSize: { widthCm: 21.59, heightCm: 27.94 },
        margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 },
      },
      "thisSection",
    );
    const merged = deleteBackward(changed.model, changed.selection);
    expect(merged.model.sections).toHaveLength(1);
    expect(sectionPageSize(merged.model, 0).widthCm).toBeCloseTo(14.8, 3);
  });
});

describe("setPageSetup apply-to scope", () => {
  const A5 = { pageSize: { widthCm: 14.8, heightCm: 21 }, margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 } };

  it("'thisSection' sets only the caret's section", () => {
    const split = insertSectionBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "nextPage");
    const { model } = setPageSetup(split.model, split.selection, A5, "thisSection");
    expect(sectionPageSize(model, 0).heightCm).toBeCloseTo(29.7, 2);
    expect(sectionPageSize(model, 1).heightCm).toBeCloseTo(21, 2);
    expect(sectionMargins(model, 1).topCm).toBeCloseTo(2, 3);
    expect(model.defaultStyles.pageSize.heightCm).toBeCloseTo(29.7, 2);
  });

  it("'wholeDocument' sets the defaults and drops per-section overrides", () => {
    const split = insertSectionBreak(fixture("HelloWorld"), caretSelection(caret(0, 5)), "nextPage");
    const scoped = setPageSetup(split.model, split.selection, A5, "thisSection");
    const { model } = setPageSetup(scoped.model, scoped.selection, {
      pageSize: { widthCm: 21.59, heightCm: 27.94 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
    });
    expect(model.sections[0].pageSize).toBeUndefined();
    expect(model.sections[1].pageSize).toBeUndefined();
    expect(sectionPageSize(model, 0).heightCm).toBeCloseTo(27.94, 2);
    expect(sectionPageSize(model, 1).heightCm).toBeCloseTo(27.94, 2);
  });

  it("'thisSection' clears the section's preserved sectPr", () => {
    const model = fixture("HelloWorld");
    const split = insertSectionBreak(model, caretSelection(caret(0, 5)), "nextPage");
    split.model.sections[1].docxSectPr = "<w:sectPr/>";
    const { model: next } = setPageSetup(split.model, split.selection, A5, "thisSection");
    expect(next.sections[1].docxSectPr).toBeUndefined();
  });
});

describe("paragraph pagination flags", () => {
  it("keepNext / keepLines / pageBreakBefore set and clear", () => {
    const on = applyParagraphFormat(fixture("Text"), caretSelection(caret(0, 0)), {
      keepNext: true,
      keepLines: true,
      pageBreakBefore: true,
    });
    const first = on.model.sections[0].blocks[0] as Paragraph;
    expect(first.formatting).toEqual({ keepNext: true, keepLines: true, pageBreakBefore: true });
    const off = applyParagraphFormat(on.model, caretSelection(caret(0, 0)), {
      keepNext: false,
      keepLines: false,
      pageBreakBefore: false,
    });
    expect((off.model.sections[0].blocks[0] as Paragraph).formatting).toBeUndefined();
  });

  it("widow control is on by default and stores only the explicit off", () => {
    const model = fixture("Text");
    expect(paragraphPaginationState(model, caretSelection(caret(0, 0))).widowControl).toBe(true);
    const off = applyParagraphFormat(model, caretSelection(caret(0, 0)), { widowControl: false });
    expect((off.model.sections[0].blocks[0] as Paragraph).formatting).toEqual({ widowControl: false });
    expect(paragraphPaginationState(off.model, caretSelection(caret(0, 0))).widowControl).toBe(false);
    const on = applyParagraphFormat(off.model, caretSelection(caret(0, 0)), { widowControl: true });
    expect((on.model.sections[0].blocks[0] as Paragraph).formatting).toBeUndefined();
  });

  it("paginationState reflects the first touched paragraph", () => {
    const model = fixture("A", "B");
    const on = applyParagraphFormat(model, caretSelection(caret(1, 0)), { keepNext: true });
    expect(paragraphPaginationState(on.model, caretSelection(caret(1, 0)))).toEqual({
      keepNext: true,
      keepLines: false,
      pageBreakBefore: false,
      widowControl: true,
    });
  });
});
