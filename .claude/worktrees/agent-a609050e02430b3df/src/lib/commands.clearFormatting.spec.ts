import { describe, expect, it } from "vitest";
import {
  applyPainterFormat,
  applyParagraphStyle,
  caretSelection,
  capturePainterFormat,
  clearFormatting,
  isWholeParagraphSelected,
  type ModelSelection,
} from "./commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import type { CaretPosition } from "./textEditing";

// SPEC-020 — clear formatting and the format painter. Clear formatting is
// spot-checked against the same layered-formatting machinery SPEC-019 pins
// (applyParagraphStyle/resetParagraphStyle): it is "reset to Normal" with the
// direct-formatting layer forced empty rather than preserved. The format
// painter commands are pure copy/paste of formatting values, so the tests
// pin the "replace, don't merge" contract that distinguishes them from
// setFormat/applyFormat.

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function caret(block: number, offset: number): CaretPosition {
  return { section: 0, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function blocks(model: DocumentModel): Paragraph[] {
  return model.sections[0].blocks.map((block) => asParagraph(block));
}

function fixture(): DocumentModel {
  const model = createDocument("clear-formatting-commands");
  const listItem = paragraph(["List item", {}]);
  listItem.list = { type: "bullet", level: 1 };
  model.sections = [
    {
      blocks: [
        paragraph(["First ", { bold: true }], ["paragraph", { italic: true, color: "#ff0000" }]),
        paragraph(["Second paragraph", {}]),
        listItem,
      ],
    },
  ];
  return model;
}

describe("clearFormatting (AC-1, AC-2, AC-4, AC-6)", () => {
  it("strips every run's character formatting across the selection's paragraphs", () => {
    const model = fixture();
    const { model: next } = clearFormatting(model, range(caret(0, 0), caret(0, 5)));
    const first = blocks(next)[0];
    expect(first.runs[0].formatting).toEqual({});
    expect(first.runs[1].formatting).toEqual({});
  });

  it("resets the touched paragraph's style to Normal, dropping heading level and paragraph formatting", () => {
    const model = fixture();
    const styled = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1").model;
    expect(blocks(styled)[0].styleId).toBe("Heading1");
    expect(blocks(styled)[0].headingLevel).toBe(1);

    const { model: next } = clearFormatting(styled, caretSelection(caret(0, 0)));
    const first = blocks(next)[0];
    expect(first.styleId).toBe("Normal");
    expect(first.headingLevel).toBeUndefined();
    expect(first.formatting).toBeUndefined();
    // The style's own bold/size disappear along with the direct formatting.
    expect(first.runs[0].formatting).toEqual({});
  });

  it("does not remove text content", () => {
    const model = fixture();
    const { model: next } = clearFormatting(model, range(caret(0, 0), caret(0, 15)));
    const first = blocks(next)[0];
    expect(first.runs.map((run) => run.text).join("")).toBe("First paragraph");
  });

  it("does not remove list membership", () => {
    const model = fixture();
    const { model: next } = clearFormatting(model, caretSelection(caret(2, 0)));
    expect(blocks(next)[2].list).toEqual({ type: "bullet", level: 1 });
  });

  it("leaves paragraphs outside the selection untouched, keeping their identity", () => {
    const model = fixture();
    const { model: next } = clearFormatting(model, caretSelection(caret(0, 0)));
    expect(blocks(next)[1]).toBe(blocks(model)[1]);
  });

  it("resolves the document's own Normal definition over the built-in", () => {
    const model = fixture();
    model.styles = [{ id: "Normal", name: "Normal", runFormatting: { fontFamily: "Georgia" } }];
    const { model: next } = clearFormatting(model, range(caret(0, 0), caret(0, 5)));
    expect(blocks(next)[0].runs[0].formatting).toEqual({ fontFamily: "Georgia" });
  });

  it("is a no-op query away from the selection's other paragraphs (AC-2's 'selected paragraphs')", () => {
    const model = fixture();
    const { model: next } = clearFormatting(model, caretSelection(caret(1, 0)));
    // Paragraph 0, never touched by a caret in paragraph 1, keeps its
    // original direct formatting.
    expect(blocks(next)[0].runs[0].formatting).toEqual({ bold: true });
  });
});

describe("isWholeParagraphSelected (AC-10)", () => {
  it("is true for a selection spanning a paragraph's exact start to end", () => {
    const model = fixture();
    expect(isWholeParagraphSelected(model, range(caret(1, 0), caret(1, 16)))).toBe(true);
  });

  it("is false for a partial range within a paragraph", () => {
    const model = fixture();
    expect(isWholeParagraphSelected(model, range(caret(1, 0), caret(1, 6)))).toBe(false);
  });

  it("is false for a range spanning more than one paragraph", () => {
    const model = fixture();
    expect(isWholeParagraphSelected(model, range(caret(0, 0), caret(1, 16)))).toBe(false);
  });

  it("is true for a collapsed caret in an otherwise-empty paragraph", () => {
    const model = fixture();
    model.sections[0].blocks[1] = paragraph();
    expect(isWholeParagraphSelected(model, caretSelection(caret(1, 0)))).toBe(true);
  });
});

describe("capturePainterFormat (AC-9, AC-10)", () => {
  it("captures only character formatting for a collapsed caret", () => {
    const model = fixture();
    const painter = capturePainterFormat(model, caretSelection(caret(0, 1)));
    expect(painter.run).toEqual({ bold: true });
    expect(painter.paragraph).toBeUndefined();
  });

  it("also captures paragraph formatting when a whole paragraph is selected", () => {
    const model = fixture();
    const styled = applyParagraphStyle(
      model,
      caretSelection(caret(1, 0)),
      "Heading1",
    ).model;
    styled.sections[0].blocks[1] = {
      ...asParagraph(styled.sections[0].blocks[1]),
      formatting: { alignment: "center" },
    };
    const painter = capturePainterFormat(styled, range(caret(1, 0), caret(1, 16)));
    expect(painter.paragraph).toEqual({ alignment: "center" });
  });

  it("omits paragraph formatting for a partial selection", () => {
    const model = fixture();
    const painter = capturePainterFormat(model, range(caret(0, 0), caret(0, 5)));
    expect(painter.paragraph).toBeUndefined();
  });
});

describe("applyPainterFormat (AC-12, AC-19)", () => {
  it("replaces (not merges into) the target range's character formatting", () => {
    const model = fixture();
    const { model: next } = applyPainterFormat(model, range(caret(1, 0), caret(1, 6)), {
      run: { bold: true },
    });
    const second = blocks(next)[1];
    expect(second.runs[0].formatting).toEqual({ bold: true });
  });

  it("replaces the earlier direct formatting entirely, not just adds to it", () => {
    const model = fixture();
    const { model: next } = applyPainterFormat(model, range(caret(0, 6), caret(0, 15)), {
      run: { underline: true },
    });
    const first = blocks(next)[0];
    // Was { italic: true, color: "#ff0000" } — both gone, only underline remains.
    expect(first.runs[1].formatting).toEqual({ underline: true });
  });

  it("does nothing to characters for a collapsed selection", () => {
    const model = fixture();
    const { model: next } = applyPainterFormat(model, caretSelection(caret(1, 0)), {
      run: { bold: true },
    });
    expect(blocks(next)[1].runs[0].formatting).toEqual({});
  });

  it("replaces every touched paragraph's own formatting when paragraph formatting was copied", () => {
    const model = fixture();
    model.sections[0].blocks[1] = {
      ...asParagraph(model.sections[0].blocks[1]),
      formatting: { alignment: "left", spaceBeforePt: 12 },
    };
    const { model: next } = applyPainterFormat(model, caretSelection(caret(1, 0)), {
      run: {},
      paragraph: { alignment: "center" },
    });
    expect(blocks(next)[1].formatting).toEqual({ alignment: "center" });
  });

  it("leaves paragraph formatting alone when none was copied", () => {
    const model = fixture();
    model.sections[0].blocks[1] = {
      ...asParagraph(model.sections[0].blocks[1]),
      formatting: { alignment: "left" },
    };
    const { model: next } = applyPainterFormat(model, range(caret(1, 0), caret(1, 6)), {
      run: { bold: true },
    });
    expect(blocks(next)[1].formatting).toEqual({ alignment: "left" });
  });
});
