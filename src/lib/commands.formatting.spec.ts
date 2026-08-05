// SPEC-013 — the value-format command layer: applying font family/size,
// colour, highlight and sub/superscript across a selection, the toolbar's
// value/mixed queries, and the pending-format patch semantics insertText
// applies at a collapsed caret.
import { describe, it, expect } from "vitest";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { applyFormatPatch, type CaretPosition } from "./textEditing";
import {
  applyFormat,
  caretSelection,
  formatValue,
  insertText,
  toggleVerticalAlign,
  verticalAlignState,
  type ModelSelection,
} from "./commands";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function doc(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("commands-formatting-test");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

describe("applyFormatPatch (pending-format semantics)", () => {
  it("sets value fields", () => {
    expect(applyFormatPatch({}, { fontFamily: "Georgia", fontSizePt: 14, color: "#ff0000" }))
      .toEqual({ fontFamily: "Georgia", fontSizePt: 14, color: "#ff0000" });
  });

  it("stores highlight false — 'No colour' is explicit, not an absence", () => {
    expect(applyFormatPatch({ highlight: "yellow" }, { highlight: false }))
      .toEqual({ highlight: false });
  });

  it("deletes vertAlign on false, so the run merges with plain runs", () => {
    expect(applyFormatPatch({ vertAlign: "subscript", bold: true }, { vertAlign: false }))
      .toEqual({ bold: true });
  });

  it("keeps SPEC-012's boolean semantics: false deletes the key", () => {
    expect(applyFormatPatch({ bold: true, italic: true }, { bold: false }))
      .toEqual({ italic: true });
  });
});

describe("applyFormat", () => {
  it("applies a value patch to exactly the selection, splitting runs", () => {
    const model = doc(paragraph(["Hello world"]));
    const { model: next, selection } = applyFormat(
      model,
      range(caret(0, 6), caret(0, 11)),
      { fontFamily: "Georgia", fontSizePt: 18 },
    );
    expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
      { text: "Hello ", formatting: {} },
      { text: "world", formatting: { fontFamily: "Georgia", fontSizePt: 18 } },
    ]);
    // The selection survives, like setFormat's (SPEC-012 AC-13/14).
    expect(selection).toEqual({ anchor: caret(0, 6), focus: caret(0, 11) });
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph(["abc"]));
    const before = JSON.stringify(model);
    applyFormat(model, range(caret(0, 0), caret(0, 3)), { color: "#00ff00" });
    expect(JSON.stringify(model)).toBe(before);
  });

  it("is a no-op for a collapsed selection (pending format is the caller's)", () => {
    const model = doc(paragraph(["abc"]));
    const { model: next } = applyFormat(model, caretSelection(caret(0, 1)), { fontSizePt: 30 });
    expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([{ text: "abc", formatting: {} }]);
  });

  it("sets an explicit no-colour highlight that does not merge away", () => {
    const model = doc(paragraph(["ab", { highlight: "yellow" }], ["cd"]));
    const { model: next } = applyFormat(model, range(caret(0, 0), caret(0, 4)), {
      highlight: false,
    });
    expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
      { text: "abcd", formatting: { highlight: false } },
    ]);
  });
});

describe("formatValue", () => {
  it("reports the caret run's value for a collapsed selection", () => {
    const model = doc(paragraph(["abc", { fontFamily: "Georgia" }]));
    expect(formatValue(model, caretSelection(caret(0, 2)), "fontFamily"))
      .toEqual({ mixed: false, value: "Georgia" });
  });

  it("prefers the pending format at a collapsed caret", () => {
    const model = doc(paragraph(["abc", { fontSizePt: 11 }]));
    expect(formatValue(model, caretSelection(caret(0, 2)), "fontSizePt", { fontSizePt: 24 }))
      .toEqual({ mixed: false, value: 24 });
  });

  it("reports the shared value over a uniform selection", () => {
    const model = doc(paragraph(["ab", { color: "#ff0000" }], ["cd", { color: "#ff0000" }]));
    expect(formatValue(model, range(caret(0, 0), caret(0, 4)), "color"))
      .toEqual({ mixed: false, value: "#ff0000" });
  });

  it("reports mixed when the selection's runs disagree (blank control)", () => {
    const model = doc(paragraph(["ab", { fontSizePt: 11 }], ["cd", { fontSizePt: 14 }]));
    expect(formatValue(model, range(caret(0, 0), caret(0, 4)), "fontSizePt"))
      .toEqual({ mixed: true });
  });

  it("reports mixed between a set and an unset run", () => {
    const model = doc(paragraph(["ab", { fontFamily: "Arial" }], ["cd"]));
    expect(formatValue(model, range(caret(0, 0), caret(0, 4)), "fontFamily"))
      .toEqual({ mixed: true });
  });

  it("treats highlight false and unset as the same 'no colour' state", () => {
    const model = doc(paragraph(["ab", { highlight: false }], ["cd"]));
    expect(formatValue(model, range(caret(0, 0), caret(0, 4)), "highlight"))
      .toEqual({ mixed: false, value: false });
  });

  it("reports mixed between a highlight and no highlight", () => {
    const model = doc(paragraph(["ab", { highlight: "yellow" }], ["cd"]));
    expect(formatValue(model, range(caret(0, 0), caret(0, 4)), "highlight"))
      .toEqual({ mixed: true });
  });
});

describe("verticalAlignState / toggleVerticalAlign", () => {
  it("tracks on/off/mixed like the boolean toggles", () => {
    const model = doc(paragraph(["ab", { vertAlign: "subscript" }], ["cd"]));
    expect(verticalAlignState(model, range(caret(0, 0), caret(0, 2)), "subscript")).toBe("on");
    expect(verticalAlignState(model, range(caret(0, 2), caret(0, 4)), "subscript")).toBe("off");
    expect(verticalAlignState(model, range(caret(0, 0), caret(0, 4)), "subscript")).toBe("mixed");
    expect(verticalAlignState(model, range(caret(0, 0), caret(0, 2)), "superscript")).toBe("off");
  });

  it("honours the pending format at a collapsed caret", () => {
    const model = doc(paragraph(["abc"]));
    expect(
      verticalAlignState(model, caretSelection(caret(0, 1)), "superscript", {
        vertAlign: "superscript",
      }),
    ).toBe("on");
  });

  it("toggles on over a mixed selection, then off", () => {
    const model = doc(paragraph(["ab", { vertAlign: "subscript" }], ["cd"]));
    const selection = range(caret(0, 0), caret(0, 4));
    const on = toggleVerticalAlign(model, selection, "subscript");
    expect(asParagraph(on.model.sections[0].blocks[0]).runs).toEqual([
      { text: "abcd", formatting: { vertAlign: "subscript" } },
    ]);
    const off = toggleVerticalAlign(on.model, on.selection, "subscript");
    expect(asParagraph(off.model.sections[0].blocks[0]).runs).toEqual([{ text: "abcd", formatting: {} }]);
  });

  it("activating one deactivates the other — mutually exclusive", () => {
    const model = doc(paragraph(["abcd", { vertAlign: "subscript" }]));
    const selection = range(caret(0, 0), caret(0, 4));
    const { model: next } = toggleVerticalAlign(model, selection, "superscript");
    expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
      { text: "abcd", formatting: { vertAlign: "superscript" } },
    ]);
  });
});

describe("insertText with a value pending format", () => {
  it("typed text carries the pending font/colour", () => {
    const model = doc(paragraph(["ab"]));
    const { model: next } = insertText(model, caretSelection(caret(0, 2)), "X", {
      fontFamily: "Georgia",
      color: "#ff0000",
    });
    expect(asParagraph(next.sections[0].blocks[0]).runs).toEqual([
      { text: "ab", formatting: {} },
      { text: "X", formatting: { fontFamily: "Georgia", color: "#ff0000" } },
    ]);
  });

  it("typed text carries a pending subscript, and false clears it", () => {
    const model = doc(paragraph(["ab", { vertAlign: "superscript" }]));
    const { model: next } = insertText(model, caretSelection(caret(0, 2)), "X", {
      vertAlign: "subscript",
    });
    expect(asParagraph(next.sections[0].blocks[0]).runs[1]).toEqual({
      text: "X",
      formatting: { vertAlign: "subscript" },
    });
    const { model: cleared } = insertText(model, caretSelection(caret(0, 2)), "X", {
      vertAlign: false,
    });
    expect(asParagraph(cleared.sections[0].blocks[0]).runs[1]).toEqual({ text: "X", formatting: {} });
  });
});
