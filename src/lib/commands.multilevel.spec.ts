// SPEC-032 — the multilevel-list and numbering-control commands, verified
// through the derived markers (listMarkers) rather than raw ids: the ids
// are minted, so behaviour is what matters.
import { describe, expect, it } from "vitest";
import {
  applyListDefinition,
  applyMultilevelPreset,
  applyParagraphStyle,
  caretSelection,
  continueListNumbering,
  linkHeadingNumbering,
  restartListNumbering,
  selectionListDefinitionLevels,
  selectionTabStops,
  setListNumberingValue,
  setTabStops,
  type ModelSelection,
} from "./commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { cmToTwips } from "./docx/units";
import { listMarkers } from "./lists";
import type { CaretPosition } from "./textEditing";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function listItem(text: string, level = 1): Paragraph {
  const p = paragraph(text);
  p.list = { type: "number", level };
  return p;
}

function caret(block: number, offset = 0): ModelSelection {
  return caretSelection({ section: 0, block, offset } as CaretPosition);
}

function range(anchorBlock: number, focusBlock: number): ModelSelection {
  return {
    anchor: { section: 0, block: anchorBlock, offset: 0 },
    focus: { section: 0, block: focusBlock, offset: 0 },
  };
}

function docWith(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("multilevel-fixture");
  model.sections = [{ blocks }];
  return model;
}

function markers(model: DocumentModel): string[] {
  return listMarkers(model)[0];
}

describe("applyMultilevelPreset (AC-1..4)", () => {
  it("numbers a selection of plain paragraphs 1./2./3.", () => {
    const model = docWith(paragraph("a"), paragraph("b"), paragraph("c"));
    const { model: next } = applyMultilevelPreset(model, range(0, 2), "decimal");
    expect(markers(next)).toEqual(["1.", "2.", "3."]);
    expect(next.listDefinitions).toHaveLength(1);
  });

  it("keeps existing levels when applied over a list", () => {
    const model = docWith(listItem("a", 1), listItem("b", 2));
    const { model: next } = applyMultilevelPreset(model, range(0, 1), "decimal");
    expect(markers(next)).toEqual(["1.", "1.1."]);
  });

  it("two lists from the same preset get independent definitions (AC-18)", () => {
    const base = docWith(paragraph("a"), paragraph("x"), paragraph("b"));
    const first = applyMultilevelPreset(base, caret(0), "decimal").model;
    const second = applyMultilevelPreset(first, caret(2), "decimal").model;
    expect(second.listDefinitions).toHaveLength(2);
    const [a, b] = [asParagraph(second.sections[0].blocks[0]), asParagraph(second.sections[0].blocks[2])];
    expect(a.list?.defId).not.toBe(b.list?.defId);
  });

  it("never mutates the input model", () => {
    const model = docWith(paragraph("a"));
    const before = structuredClone(model);
    applyMultilevelPreset(model, caret(0), "legal");
    expect(model).toEqual(before);
  });
});

describe("applyListDefinition (AC-5..12, AC-17, AC-18)", () => {
  it("editing a defined list's levels re-renders every item of that list", () => {
    const base = docWith(paragraph("a"), paragraph("b"));
    const applied = applyMultilevelPreset(base, range(0, 1), "decimal").model;
    const edited = applyListDefinition(applied, caret(0), [
      { style: "upperRoman", prefix: "(", suffix: ")", startAt: 5 },
    ]).model;
    expect(markers(edited)).toEqual(["(V)", "(VI)"]);
    // Same definition updated in place, not a second one.
    expect(edited.listDefinitions).toHaveLength(1);
  });

  it("does not touch other lists (AC-18)", () => {
    const base = docWith(paragraph("a"), paragraph("x"), paragraph("b"));
    const first = applyMultilevelPreset(base, caret(0), "decimal").model;
    const second = applyMultilevelPreset(first, caret(2), "decimal").model;
    const edited = applyListDefinition(second, caret(0), [{ style: "upperRoman", suffix: "." }]).model;
    expect(markers(edited)).toEqual(["I.", "", "1."]);
  });

  it("switching a level across the bullet/number line updates item types", () => {
    const base = docWith(paragraph("a"));
    const applied = applyMultilevelPreset(base, caret(0), "decimal").model;
    const edited = applyListDefinition(applied, caret(0), [{ style: "disc" }]).model;
    expect(asParagraph(edited.sections[0].blocks[0]).list?.type).toBe("bullet");
    expect(markers(edited)).toEqual(["•"]);
  });

  it("selectionListDefinitionLevels returns the definition being edited", () => {
    const base = docWith(paragraph("a"));
    const applied = applyMultilevelPreset(base, caret(0), "legal").model;
    const levels = selectionListDefinitionLevels(applied, caret(0));
    expect(levels[2].includeLevels).toBe(3);
    expect(levels[2].suffix).toBeUndefined();
  });
});

describe("restart / continue / set value (AC-14..16)", () => {
  function threeItemList(): DocumentModel {
    const base = docWith(listItem("one"), listItem("two"), listItem("three"));
    return applyMultilevelPreset(base, range(0, 2), "decimal").model;
  }

  it("restart resets the clicked item and the rest of the list", () => {
    const restarted = restartListNumbering(threeItemList(), caret(1)).model;
    expect(markers(restarted)).toEqual(["1.", "1.", "2."]);
  });

  it("continue rejoins the previous sequence", () => {
    const restarted = restartListNumbering(threeItemList(), caret(1)).model;
    const continued = continueListNumbering(restarted, caret(1)).model;
    expect(markers(continued)).toEqual(["1.", "2.", "3."]);
  });

  it("set numbering value restarts at the explicit ordinal", () => {
    const set = setListNumberingValue(threeItemList(), caret(1), 42).model;
    expect(markers(set)).toEqual(["1.", "42.", "43."]);
  });

  it("works on an ad-hoc SPEC-018 list by upgrading it", () => {
    const model = docWith(listItem("one"), listItem("two"), listItem("three"));
    const restarted = restartListNumbering(model, caret(1)).model;
    expect(markers(restarted)).toEqual(["1.", "1.", "2."]);
    expect(restarted.listDefinitions?.length).toBe(1);
  });

  it("continue joins two separate ad-hoc lists across an interruption", () => {
    const model = docWith(listItem("one"), listItem("two"), paragraph("gap"), listItem("three"));
    const continued = continueListNumbering(model, caret(3)).model;
    expect(markers(continued)).toEqual(["1.", "2.", "", "3."]);
  });

  it("continue is a no-op when nothing numbered precedes", () => {
    const model = docWith(listItem("only"));
    const result = continueListNumbering(model, caret(0)).model;
    expect(markers(result)).toEqual(["1."]);
  });
});

describe("heading numbering (AC-13, AC-29..31)", () => {
  function headingDoc(): DocumentModel {
    const model = docWith(paragraph("Chapter One"), paragraph("Section A"), paragraph("Chapter Two"));
    let next = applyParagraphStyle(model, caret(0), "Heading1").model;
    next = applyParagraphStyle(next, caret(1), "Heading2").model;
    next = applyParagraphStyle(next, caret(2), "Heading1").model;
    return next;
  }

  it("linking numbers existing Heading 1-3 paragraphs", () => {
    const linked = linkHeadingNumbering(headingDoc(), caret(0)).model;
    expect(markers(linked)).toEqual(["1.", "1.1.", "2."]);
  });

  it("a newly styled heading auto-numbers through applyParagraphStyle", () => {
    let model = linkHeadingNumbering(headingDoc(), caret(0)).model;
    model.sections[0].blocks.push(paragraph("Chapter Three"));
    model = applyParagraphStyle(model, caret(3), "Heading1").model;
    expect(markers(model)).toEqual(["1.", "1.1.", "2.", "3."]);
  });

  it("deleting a heading renumbers the rest (AC-29)", () => {
    const linked = linkHeadingNumbering(headingDoc(), caret(0)).model;
    const blocks = [...linked.sections[0].blocks];
    blocks.splice(0, 1);
    const withoutFirst: DocumentModel = { ...linked, sections: [{ blocks }] };
    // The leading Heading 2 implies a phantom level-1 "1" (Word's own
    // behaviour), so the following Heading 1 counts on as 2.
    expect(markers(withoutFirst)).toEqual(["1.1.", "2."]);
  });

  it("deleting the middle of three same-level headings closes the gap (AC-29)", () => {
    const model = docWith(paragraph("Alpha"), paragraph("Beta"), paragraph("Gamma"));
    let linked = model;
    for (let block = 0; block < 3; block++) {
      linked = applyParagraphStyle(linked, caret(block), "Heading1").model;
    }
    linked = linkHeadingNumbering(linked, caret(0)).model;
    const blocks = [...linked.sections[0].blocks];
    blocks.splice(1, 1);
    const withoutMiddle: DocumentModel = { ...linked, sections: [{ blocks }] };
    expect(markers(withoutMiddle)).toEqual(["1.", "2."]);
  });

  it("moving a heading renumbers by document order (AC-30)", () => {
    const linked = linkHeadingNumbering(headingDoc(), caret(0)).model;
    const blocks = [...linked.sections[0].blocks];
    const [last] = blocks.splice(2, 1);
    blocks.unshift(last);
    const moved: DocumentModel = { ...linked, sections: [{ blocks }] };
    expect(markers(moved)).toEqual(["1.", "2.", "2.1."]);
  });

  it("demoting Heading1 -> Heading2 renumbers at the deeper level (AC-31)", () => {
    let model = linkHeadingNumbering(headingDoc(), caret(0)).model;
    model = applyParagraphStyle(model, caret(2), "Heading2").model;
    expect(markers(model)).toEqual(["1.", "1.1.", "1.2."]);
  });

  it("applying an unlinked style removes the linked numbering", () => {
    let model = linkHeadingNumbering(headingDoc(), caret(0)).model;
    model = applyParagraphStyle(model, caret(0), "Normal").model;
    expect(asParagraph(model.sections[0].blocks[0]).list).toBeUndefined();
  });
});

describe("setTabStops (AC-19..25)", () => {
  it("stores snapped, sorted, deduplicated stops on touched paragraphs", () => {
    const model = docWith(paragraph("a"));
    const { model: next } = setTabStops(model, caret(0), [
      { positionCm: 8, alignment: "right", leader: "dot" },
      { positionCm: 2, alignment: "left" },
      { positionCm: 2, alignment: "decimal" },
    ]);
    const stops = asParagraph(next.sections[0].blocks[0]).formatting?.tabStops;
    // Positions snap to the shared twip grid (like every length command),
    // so they are compared in twips rather than raw centimetres.
    expect(stops?.map((stop) => [cmToTwips(stop.positionCm), stop.alignment, stop.leader])).toEqual([
      [cmToTwips(2), "decimal", undefined],
      [cmToTwips(8), "right", "dot"],
    ]);
    expect(selectionTabStops(next, caret(0))).toEqual(stops);
  });

  it("an empty list removes the field", () => {
    const model = docWith(paragraph("a"));
    const withStops = setTabStops(model, caret(0), [{ positionCm: 3, alignment: "center" }]).model;
    const cleared = setTabStops(withStops, caret(0), []).model;
    expect(asParagraph(cleared.sections[0].blocks[0]).formatting?.tabStops).toBeUndefined();
  });
});
