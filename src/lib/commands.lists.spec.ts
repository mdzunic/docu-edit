import { describe, expect, it } from "vitest";
import {
  applyListAutoformat,
  caretSelection,
  changeListLevel,
  removeListFormat,
  setListMarker,
  splitBlock,
  toggleList,
  toolbarListState,
  type ModelSelection,
} from "./commands";
import { createDocument, type DocumentModel, type ListFormatting, type Paragraph } from "./documentModel";
import type { CaretPosition } from "./textEditing";

function paragraph(text: string, list?: ListFormatting): Paragraph {
  const p: Paragraph = {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: {} }] : [],
  };
  if (list) {
    p.list = list;
  }
  return p;
}

function docWith(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("commands-lists-fixture");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset = 0): CaretPosition {
  return { section: 0, block, offset };
}

function range(anchorBlock: number, focusBlock: number): ModelSelection {
  return { anchor: caret(anchorBlock), focus: caret(focusBlock) };
}

function lists(model: DocumentModel): Array<ListFormatting | undefined> {
  return model.sections[0].blocks.map((block) => block.list);
}

describe("SPEC-018 AC-1/AC-2 — toggleList", () => {
  it("applies list formatting at level 1 to plain paragraphs", () => {
    const model = docWith(paragraph("a"), paragraph("b"));
    const { model: next } = toggleList(model, range(0, 1), "bullet");
    expect(lists(next)).toEqual([
      { type: "bullet", level: 1 },
      { type: "bullet", level: 1 },
    ]);
  });

  it("removes list formatting when every touched paragraph already has the type", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 1 }),
      paragraph("b", { type: "bullet", level: 3 }),
    );
    const { model: next } = toggleList(model, range(0, 1), "bullet");
    expect(lists(next)).toEqual([undefined, undefined]);
  });

  it("applies to all when only some are of the type, keeping existing levels", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 2 }),
      paragraph("b"),
    );
    const { model: next } = toggleList(model, range(0, 1), "bullet");
    expect(lists(next)).toEqual([
      { type: "bullet", level: 2 },
      { type: "bullet", level: 1 },
    ]);
  });

  it("converts items of the other type in place, dropping their stale marker", () => {
    const model = docWith(
      paragraph("a", { type: "number", level: 2, marker: "upperRoman", startAt: 4 }),
      paragraph("b"),
    );
    const { model: next } = toggleList(model, range(0, 1), "bullet");
    expect(lists(next)).toEqual([
      { type: "bullet", level: 2, startAt: 4 },
      { type: "bullet", level: 1 },
    ]);
  });

  it("does not mutate the input model", () => {
    const model = docWith(paragraph("a"));
    toggleList(model, caretSelection(caret(0)), "number");
    expect(model.sections[0].blocks[0].list).toBeUndefined();
  });
});

describe("SPEC-018 AC-3/AC-4 — changeListLevel", () => {
  it("demotes and promotes touched list items by one level", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 1 }),
      paragraph("b", { type: "number", level: 5 }),
    );
    const demoted = changeListLevel(model, range(0, 1), 1).model;
    expect(lists(demoted).map((l) => l?.level)).toEqual([2, 6]);
    const promoted = changeListLevel(model, range(0, 1), -1).model;
    expect(lists(promoted).map((l) => l?.level)).toEqual([1, 4]);
  });

  it("clamps at level 9 and level 1 (no effect at the bounds)", () => {
    const model = docWith(
      paragraph("max", { type: "bullet", level: 9 }),
      paragraph("min", { type: "bullet", level: 1 }),
    );
    const demoted = changeListLevel(model, caretSelection(caret(0, 0)), 1).model;
    expect(demoted.sections[0].blocks[0]).toBe(model.sections[0].blocks[0]);
    const promoted = changeListLevel(model, caretSelection(caret(1, 0)), -1).model;
    expect(promoted.sections[0].blocks[1]).toBe(model.sections[0].blocks[1]);
  });

  it("leaves non-list paragraphs in the selection untouched", () => {
    const model = docWith(paragraph("plain"), paragraph("item", { type: "bullet", level: 1 }));
    const { model: next } = changeListLevel(model, range(0, 1), 1);
    expect(next.sections[0].blocks[0]).toBe(model.sections[0].blocks[0]);
    expect(next.sections[0].blocks[1].list?.level).toBe(2);
  });

  it("drops an imported per-level indent when the level changes", () => {
    const model = docWith(paragraph("a", { type: "bullet", level: 1, indentCm: 3 }));
    const { model: next } = changeListLevel(model, caretSelection(caret(0)), 1);
    expect(next.sections[0].blocks[0].list).toEqual({ type: "bullet", level: 2 });
  });
});

describe("SPEC-018 AC-12/AC-13 — setListMarker", () => {
  it("applies the marker to every item at the current level of the current list", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 1 }),
      paragraph("b", { type: "bullet", level: 2 }),
      paragraph("c", { type: "bullet", level: 1 }),
      paragraph("plain"),
      paragraph("other list", { type: "bullet", level: 1 }),
    );
    const { model: next } = setListMarker(model, caretSelection(caret(0)), "square");
    expect(lists(next)).toEqual([
      { type: "bullet", level: 1, marker: "square" },
      { type: "bullet", level: 2 }, // other level untouched
      { type: "bullet", level: 1, marker: "square" },
      undefined,
      { type: "bullet", level: 1 }, // beyond the interruption: another list
    ]);
  });

  it("picking a number style converts the level's items to numbered", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 1 }),
      paragraph("b", { type: "bullet", level: 1 }),
    );
    const { model: next } = setListMarker(model, caretSelection(caret(1)), "upperAlpha");
    expect(lists(next)).toEqual([
      { type: "number", level: 1, marker: "upperAlpha" },
      { type: "number", level: 1, marker: "upperAlpha" },
    ]);
  });

  it("is a no-op outside a list", () => {
    const model = docWith(paragraph("plain"));
    const { model: next } = setListMarker(model, caretSelection(caret(0)), "square");
    expect(next.sections[0].blocks[0]).toBe(model.sections[0].blocks[0]);
  });
});

describe("SPEC-018 AC-8/AC-9 — removeListFormat", () => {
  it("clears list formatting without touching text", () => {
    const model = docWith(paragraph("keep me", { type: "number", level: 3 }));
    const { model: next } = removeListFormat(model, caretSelection(caret(0)));
    expect(next.sections[0].blocks[0].list).toBeUndefined();
    expect(next.sections[0].blocks[0].runs[0].text).toBe("keep me");
  });
});

describe("SPEC-018 AC-6 — Enter creates the next item at the same level", () => {
  it("splitBlock carries the list onto the new paragraph", () => {
    const model = docWith(paragraph("item text", { type: "bullet", level: 2, marker: "square" }));
    const { model: next, selection } = splitBlock(model, caretSelection(caret(0, 4)));
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(2);
    expect(blocks[0].list).toEqual({ type: "bullet", level: 2, marker: "square" });
    expect(blocks[1].list).toEqual({ type: "bullet", level: 2, marker: "square" });
    // Independent objects: editing one item's list must not alias the other.
    expect(blocks[1].list).not.toBe(blocks[0].list);
    expect(selection.focus).toEqual({ section: 0, block: 1, offset: 0 });
  });
});

describe("SPEC-018 AC-14..16 — applyListAutoformat", () => {
  it("replaces the prefix paragraph with an empty level-1 item", () => {
    const model = docWith(paragraph("1. "));
    const { model: next, selection } = applyListAutoformat(
      model,
      caretSelection(caret(0, 3)),
      "number",
    );
    const block = next.sections[0].blocks[0];
    expect(block.runs).toEqual([]);
    expect(block.list).toEqual({ type: "number", level: 1 });
    expect(selection.focus).toEqual({ section: 0, block: 0, offset: 0 });
    // The input model still holds the literal prefix — the undo snapshot.
    expect(model.sections[0].blocks[0].runs[0].text).toBe("1. ");
  });
});

describe("toolbarListState", () => {
  it("reports pressed states matching the toggle predicate and the first item's marker", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 2 }),
      paragraph("b", { type: "number", level: 1 }),
    );
    expect(toolbarListState(model, caretSelection(caret(0)))).toEqual({
      bulleted: true,
      numbered: false,
      marker: "circle",
    });
    expect(toolbarListState(model, range(0, 1))).toEqual({
      bulleted: false,
      numbered: false,
      marker: "circle",
    });
    expect(toolbarListState(model, caretSelection(caret(1)))).toEqual({
      bulleted: false,
      numbered: true,
      marker: "decimal",
    });
  });
});
