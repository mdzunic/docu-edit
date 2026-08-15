import { describe, it, expect } from "vitest";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { blockText, type CaretPosition } from "./textEditing";
import {
  adjustIndent,
  applyParagraphFormat,
  blockCount,
  blockOrdinal,
  caretFormatting,
  caretSelection,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentEnd,
  documentStart,
  firstTouchedFormatting,
  formatState,
  hasSpaceBefore,
  INDENT_STEP_CM,
  insertLineBreak,
  insertText,
  isCollapsed,
  orderedRange,
  selectAll,
  setAlignment,
  setFormat,
  splitBlock,
  toggleFormat,
  toggleSpaceBefore,
  touchedBlocks,
  type ModelSelection,
} from "./commands";
import { cmToTwips, twipsToCm } from "./core/units";

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
    const runs = asParagraph(next.sections[0].blocks[0]).runs;
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
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { alignment: "center", indentStartCm: 1 },
      styleId: "Fancy",
      headingLevel: 2,
    };
    const { model: next } = splitBlock(model, caretSelection(caret(0, 8)));
    const [first, second] = next.sections[0].blocks.map(asParagraph);
    expect(first.formatting).toEqual({ alignment: "center", indentStartCm: 1 });
    expect(second.formatting).toEqual({ alignment: "center", indentStartCm: 1 });
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

describe("character formatting (SPEC-012)", () => {
  it("insertText applies a pending format override to the inserted text (AC-10)", () => {
    const model = doc(paragraph(["plain "], ["bold", { bold: true }]));
    const { model: next } = insertText(model, caretSelection(caret(0, 6)), "X", { bold: true });
    const runs = asParagraph(next.sections[0].blocks[0]).runs;
    expect(runs.map((r) => r.text)).toEqual(["plain ", "Xbold"]);
  });

  describe("caretFormatting", () => {
    it("is the formatting new text at the caret would inherit (AC-12)", () => {
      const model = doc(paragraph(["bold", { bold: true }], [" plain"]));
      expect(caretFormatting(model, caret(0, 4))).toEqual({ bold: true });
      expect(caretFormatting(model, caret(0, 10))).toEqual({});
    });
  });

  describe("formatState", () => {
    it("is 'off' for a collapsed caret in unformatted text", () => {
      const model = doc(paragraph(["plain"]));
      expect(formatState(model, caretSelection(caret(0, 2)), "bold")).toBe("off");
    });

    it("is 'on' for a collapsed caret inside bold text (AC-12)", () => {
      const model = doc(paragraph(["bold", { bold: true }]));
      expect(formatState(model, caretSelection(caret(0, 2)), "bold")).toBe("on");
    });

    it("a pending override wins over the run's own formatting (AC-9)", () => {
      const model = doc(paragraph(["plain"]));
      expect(formatState(model, caretSelection(caret(0, 2)), "bold", { bold: true })).toBe("on");
      const boldModel = doc(paragraph(["bold", { bold: true }]));
      expect(formatState(boldModel, caretSelection(caret(0, 2)), "bold", { bold: false })).toBe(
        "off",
      );
    });

    it("is 'on' only when every run in the selection has the format (AC-13)", () => {
      const model = doc(paragraph(["all bold", { bold: true }]));
      expect(formatState(model, range(caret(0, 0), caret(0, 8)), "bold")).toBe("on");
    });

    it("is 'off' when no run in the selection has the format", () => {
      const model = doc(paragraph(["plain text"]));
      expect(formatState(model, range(caret(0, 0), caret(0, 10)), "bold")).toBe("off");
    });

    it("is 'mixed' when only part of the selection has the format (AC-14)", () => {
      const model = doc(paragraph(["plain ", {}], ["bold", { bold: true }]));
      expect(formatState(model, range(caret(0, 0), caret(0, 10)), "bold")).toBe("mixed");
    });
  });

  describe("setFormat", () => {
    it("sets the format across the selection without collapsing it (AC-7, AC-13)", () => {
      const model = doc(paragraph(["Hello world"]));
      const { model: next, selection } = setFormat(
        model,
        range(caret(0, 0), caret(0, 5)),
        "bold",
        true,
      );
      expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
        { text: "Hello", formatting: { bold: true } },
        { text: " world", formatting: {} },
      ]);
      expect(selection).toEqual(range(caret(0, 0), caret(0, 5)));
      expect(isCollapsed(selection)).toBe(false);
    });

    it("is a no-op for a collapsed selection but still returns a fresh state", () => {
      const model = doc(paragraph(["abc"]));
      const { model: next } = setFormat(model, caretSelection(caret(0, 1)), "bold", true);
      expect(next).not.toBe(model);
      expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([{ text: "abc", formatting: {} }]);
    });

    it("never mutates the input model", () => {
      const model = doc(paragraph(["Hello world"]));
      const before = snapshot(model);
      setFormat(model, range(caret(0, 0), caret(0, 5)), "bold", true);
      expect(snapshot(model)).toBe(before);
    });
  });

  describe("toggleFormat", () => {
    it("turns the format on when the selection is off (AC-16)", () => {
      const model = doc(paragraph(["Hello"]));
      const { model: next } = toggleFormat(model, range(caret(0, 0), caret(0, 5)), "bold");
      expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
        { text: "Hello", formatting: { bold: true } },
      ]);
    });

    it("turns the format off when the whole selection already has it (AC-16)", () => {
      const model = doc(paragraph(["Hello", { bold: true }]));
      const { model: next } = toggleFormat(model, range(caret(0, 0), caret(0, 5)), "bold");
      expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([{ text: "Hello", formatting: {} }]);
    });

    it("a mixed selection's first toggle makes the whole selection bold (AC-15)", () => {
      const model = doc(paragraph(["bo", { bold: true }], ["ld and plain"]));
      const { model: next } = toggleFormat(model, range(caret(0, 0), caret(0, 14)), "bold");
      expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
        { text: "bold and plain", formatting: { bold: true } },
      ]);
    });

    it("a mixed selection's second toggle clears the format everywhere (AC-17)", () => {
      const model = doc(paragraph(["bo", { bold: true }], ["ld and plain"]));
      const once = toggleFormat(model, range(caret(0, 0), caret(0, 14)), "bold");
      const twice = toggleFormat(once.model, once.selection, "bold");
      expect(asParagraph(twice.model.sections[0].blocks[0]).runs).toEqual([
        { text: "bold and plain", formatting: {} },
      ]);
    });

    it("splits and re-merges to the minimum number of runs over a run boundary (AC-20, AC-21)", () => {
      const model = doc(paragraph(["one", { bold: true }], ["two", { bold: true }], ["three"]));
      const { model: next } = toggleFormat(model, range(caret(0, 0), caret(0, 11)), "bold");
      expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
        { text: "onetwothree", formatting: { bold: true } },
      ]);
    });
  });
});

// --- paragraph formatting (SPEC-014) ----------------------------------------

describe("touchedBlocks (SPEC-014 AC-4)", () => {
  it("a collapsed caret touches exactly its own paragraph", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]), paragraph(["three"]));
    expect(touchedBlocks(model, caretSelection(caret(1, 2)))).toEqual([
      { section: 0, block: 1 },
    ]);
  });

  it("a range touches every paragraph it intersects, in document order", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]), paragraph(["three"]));
    expect(touchedBlocks(model, range(caret(2, 1), caret(0, 2)))).toEqual([
      { section: 0, block: 0 },
      { section: 0, block: 1 },
      { section: 0, block: 2 },
    ]);
  });
});

describe("setAlignment (SPEC-014 AC-2, AC-5)", () => {
  it("applies to every paragraph the selection touches", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]), paragraph(["three"]));
    const { model: next } = setAlignment(model, range(caret(0, 1), caret(1, 1)), "center");
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.alignment).toBe("center");
    expect(asParagraph(next.sections[0].blocks[1]).formatting?.alignment).toBe("center");
    expect(asParagraph(next.sections[0].blocks[2]).formatting).toBeUndefined();
  });

  it("preserves the selection and never mutates the input", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]));
    const before = snapshot(model);
    const { model: next, selection } = setAlignment(model, range(caret(1, 1), caret(0, 1)), "right");
    expect(snapshot(model)).toBe(before);
    expect(next).not.toBe(model);
    expect(selection).toEqual(range(caret(0, 1), caret(1, 1)));
  });

  it("keeps untouched paragraphs by reference for the canvas's v-memo", () => {
    const untouched = paragraph(["stays"]);
    const model = doc(paragraph(["edited"]), untouched);
    const { model: next } = setAlignment(model, caretSelection(caret(0, 0)), "justify");
    expect(next.sections[0].blocks[1]).toBe(untouched);
  });
});

describe("adjustIndent (SPEC-014 AC-6, AC-7)", () => {
  it("adds 1.27 cm to the start indent of every touched paragraph", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]));
    const { model: next } = adjustIndent(model, range(caret(0, 0), caret(1, 1)), INDENT_STEP_CM);
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.indentStartCm).toBeCloseTo(1.27, 10);
    expect(asParagraph(next.sections[0].blocks[1]).formatting?.indentStartCm).toBeCloseTo(1.27, 10);
  });

  it("accumulates on repeat and removes the same step back down", () => {
    const model = doc(paragraph(["one"]));
    const selection = caretSelection(caret(0, 0));
    const once = adjustIndent(model, selection, INDENT_STEP_CM);
    const twice = adjustIndent(once.model, selection, INDENT_STEP_CM);
    expect(asParagraph(twice.model.sections[0].blocks[0]).formatting?.indentStartCm).toBeCloseTo(2.54, 10);
    const back = adjustIndent(twice.model, selection, -INDENT_STEP_CM);
    expect(asParagraph(back.model.sections[0].blocks[0]).formatting?.indentStartCm).toBeCloseTo(1.27, 10);
  });

  it("stops at 0: a partial indent decreases to an explicit 0, never negative", () => {
    const model = doc(paragraph(["one"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { indentStartCm: 0.5 },
    };
    const { model: next } = adjustIndent(model, caretSelection(caret(0, 0)), -INDENT_STEP_CM);
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.indentStartCm).toBe(0);
  });

  it("leaves a paragraph with no start indent untouched on decrease", () => {
    const untouched = paragraph(["one"]);
    const model = doc(untouched);
    const { model: next } = adjustIndent(model, caretSelection(caret(0, 0)), -INDENT_STEP_CM);
    expect(next.sections[0].blocks[0]).toBe(untouched);
    expect(asParagraph(next.sections[0].blocks[0]).formatting).toBeUndefined();
  });

  it("evaluates independently per touched paragraph (AC-5)", () => {
    const model = doc(paragraph(["a"]), paragraph(["b"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { indentStartCm: 1.27 },
    };
    const { model: next } = adjustIndent(model, range(caret(0, 0), caret(1, 1)), INDENT_STEP_CM);
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.indentStartCm).toBeCloseTo(2.54, 10);
    expect(asParagraph(next.sections[0].blocks[1]).formatting?.indentStartCm).toBeCloseTo(1.27, 10);
  });

  it("snaps the result to the .docx twip grid so a round trip is exact (AC-26)", () => {
    const model = doc(paragraph(["one"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { indentStartCm: 1.27 },
    };
    const selection = caretSelection(caret(0, 0));
    const twice = adjustIndent(adjustIndent(model, selection, INDENT_STEP_CM).model, selection, INDENT_STEP_CM);
    const stored = asParagraph(twice.model.sections[0].blocks[0]).formatting?.indentStartCm;
    // 3 × 1.27 cm = 2160 twips exactly; the stored double must be the twip
    // grid's own value, not a floating-point accumulation next to it.
    expect(stored).toBe(twipsToCm(cmToTwips(3.81)));
    expect(cmToTwips(stored ?? 0)).toBe(2160);
  });
});

describe("space-before toggle (SPEC-014 AC-15..21)", () => {
  it("hasSpaceBefore reads the table's label column: 0 or absent is 'Add'", () => {
    expect(hasSpaceBefore({})).toBe(false);
    expect(hasSpaceBefore({ spaceBeforePt: 0 })).toBe(false);
    expect(hasSpaceBefore({ spaceBeforePt: 6 })).toBe(true);
  });

  it("sets an explicit 12 pt when space-before is absent (AC-15)", () => {
    const model = doc(paragraph(["one"]));
    const { model: next } = toggleSpaceBefore(model, caretSelection(caret(0, 0)));
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.spaceBeforePt).toBe(12);
  });

  it("writes an explicit 0 — not an absence — when removing (AC-16, AC-19)", () => {
    const model = doc(paragraph(["one"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { spaceBeforePt: 6 },
    };
    const { model: next } = toggleSpaceBefore(model, caretSelection(caret(0, 0)));
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.spaceBeforePt).toBe(0);
    expect("spaceBeforePt" in (asParagraph(next.sections[0].blocks[0]).formatting ?? {})).toBe(true);
  });

  it("removing a custom 6 pt and adding it back yields 12, never 6 (AC-20)", () => {
    const model = doc(paragraph(["one"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { spaceBeforePt: 6 },
    };
    const selection = caretSelection(caret(0, 0));
    const removed = toggleSpaceBefore(model, selection);
    const added = toggleSpaceBefore(removed.model, selection);
    expect(asParagraph(added.model.sections[0].blocks[0]).formatting?.spaceBeforePt).toBe(12);
  });

  it("applies per paragraph strictly by the table, regardless of the label source (AC-18)", () => {
    const model = doc(paragraph(["has"]), paragraph(["none"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { spaceBeforePt: 6 },
    };
    const { model: next } = toggleSpaceBefore(model, range(caret(0, 0), caret(1, 2)));
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.spaceBeforePt).toBe(0);
    expect(asParagraph(next.sections[0].blocks[1]).formatting?.spaceBeforePt).toBe(12);
  });

  it("firstTouchedFormatting reflects the first paragraph in document order (AC-17)", () => {
    const model = doc(paragraph(["has"]), paragraph(["none"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { spaceBeforePt: 6 },
    };
    // Reversed selection: the anchor sits in the second paragraph.
    const formatting = firstTouchedFormatting(model, range(caret(1, 2), caret(0, 0)));
    expect(hasSpaceBefore(formatting)).toBe(true);
  });
});

describe("applyParagraphFormat (SPEC-014 AC-12, AC-13, AC-24)", () => {
  it("sets line spacing and space before/after on every touched paragraph", () => {
    const model = doc(paragraph(["one"]), paragraph(["two"]));
    const { model: next } = applyParagraphFormat(model, range(caret(0, 0), caret(1, 1)), {
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 8,
    });
    for (const block of next.sections[0].blocks.map(asParagraph)) {
      expect(block.formatting).toEqual({ lineSpacing: 1.5, spaceBeforePt: 6, spaceAfterPt: 8 });
    }
  });

  it("null clears a field so a preset replaces an absolute line height", () => {
    const model = doc(paragraph(["one"]));
    model.sections[0].blocks[0] = {
      ...asParagraph(model.sections[0].blocks[0]),
      formatting: { lineSpacingPt: 14, lineSpacingRule: "exact" },
    };
    const { model: next } = applyParagraphFormat(model, caretSelection(caret(0, 0)), {
      lineSpacing: 2,
      lineSpacingPt: null,
      lineSpacingRule: null,
    });
    expect(asParagraph(next.sections[0].blocks[0]).formatting).toEqual({ lineSpacing: 2 });
  });

  it("direction: rtl is stored; direction: ltr removes the field (AC-24)", () => {
    const model = doc(paragraph(["one"]));
    const rtl = applyParagraphFormat(model, caretSelection(caret(0, 0)), { direction: "rtl" });
    expect(asParagraph(rtl.model.sections[0].blocks[0]).formatting?.direction).toBe("rtl");
    const ltr = applyParagraphFormat(rtl.model, caretSelection(caret(0, 0)), { direction: "ltr" });
    expect(asParagraph(ltr.model.sections[0].blocks[0]).formatting?.direction).toBeUndefined();
  });

  it("snaps point values to the .docx twentieths grid (AC-26)", () => {
    const model = doc(paragraph(["one"]));
    const { model: next } = applyParagraphFormat(model, caretSelection(caret(0, 0)), {
      spaceBeforePt: 12.34,
      lineSpacingPt: 12.34,
    });
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.spaceBeforePt).toBe(12.35);
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.lineSpacingPt).toBe(12.35);
  });

  it("snaps a custom multiple to 20ths — exact in both formats' grids (AC-26)", () => {
    const model = doc(paragraph(["one"]));
    const { model: next } = applyParagraphFormat(model, caretSelection(caret(0, 0)), {
      lineSpacing: 1.37,
    });
    expect(asParagraph(next.sections[0].blocks[0]).formatting?.lineSpacing).toBe(1.35);
  });
});
