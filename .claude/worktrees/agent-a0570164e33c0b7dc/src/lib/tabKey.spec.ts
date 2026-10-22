import { describe, expect, it } from "vitest";
import { resolveTabAction } from "./tabKey";
import { caretSelection, INDENT_STEP_CM, type ModelSelection } from "./commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph, type Table } from "./documentModel";
import type { ListFormatting } from "./core/lists";
import type { CaretPosition } from "./textEditing";

function paragraph(text: string, list?: ListFormatting): Paragraph {
  const p: Paragraph = { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
  if (list) {
    p.list = list;
  }
  return p;
}

function doc(...blocks: Array<Paragraph | Table>): DocumentModel {
  const model = createDocument("tab-key-fixture");
  model.sections = [{ blocks }];
  return model;
}

function grid(rows: string[][]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({ cells: cells.map((text) => ({ blocks: [paragraph(text)] })) })),
    columnWidthsCm: rows[0].map(() => 5),
  };
}

function caret(block: number, offset: number, cell?: { row: number; col: number; block: number }): CaretPosition {
  const position: CaretPosition = { section: 0, block, offset };
  if (cell) {
    position.cell = cell;
  }
  return position;
}

function at(block: number, offset: number, cell?: { row: number; col: number; block: number }): ModelSelection {
  return caretSelection(caret(block, offset, cell));
}

describe("resolveTabAction — rule 1, inside a table cell (SPEC-022 AC-10..12)", () => {
  const model = doc(grid([["a", "b"], ["c", "d"]]));

  it("moves to the next cell", () => {
    const action = resolveTabAction(model, at(0, 0, { row: 0, col: 0, block: 0 }), false);
    expect(action).toEqual({ kind: "moveCaret", caret: expect.objectContaining({ cell: { row: 0, col: 1, block: 0 } }) });
  });

  it("moves to the previous cell on Shift+Tab", () => {
    const action = resolveTabAction(model, at(0, 0, { row: 0, col: 1, block: 0 }), true);
    expect(action).toEqual({ kind: "moveCaret", caret: expect.objectContaining({ cell: { row: 0, col: 0, block: 0 } }) });
  });

  it("appends a row from the last cell rather than leaving the table", () => {
    const action = resolveTabAction(model, at(0, 0, { row: 1, col: 1, block: 0 }), false);
    expect(action.kind).toBe("appendRow");
  });

  it("does nothing on Shift+Tab in the very first cell", () => {
    expect(resolveTabAction(model, at(0, 0, { row: 0, col: 0, block: 0 }), true)).toEqual({ kind: "none" });
  });

  it("beats the list and indent rules — a list inside a cell still navigates", () => {
    const listed = doc(grid([["a", "b"]]));
    const cell = asParagraph(((listed.sections[0].blocks[0] as Table).rows[0].cells[0].blocks[0]));
    cell.list = { type: "bullet", level: 1 };
    const action = resolveTabAction(listed, at(0, 0, { row: 0, col: 0, block: 0 }), false);
    expect(action.kind).toBe("moveCaret");
  });
});

describe("resolveTabAction — rules 2 and 3, at the start of a paragraph", () => {
  it("changes list level for a list item (rule 2)", () => {
    const model = doc(paragraph("item", { type: "bullet", level: 1 }));
    expect(resolveTabAction(model, at(0, 0), false)).toEqual({ kind: "listLevel", delta: 1 });
    expect(resolveTabAction(model, at(0, 0), true)).toEqual({ kind: "listLevel", delta: -1 });
  });

  it("indents a plain paragraph (rule 3)", () => {
    const model = doc(paragraph("text"));
    expect(resolveTabAction(model, at(0, 0), false)).toEqual({ kind: "indent", deltaCm: INDENT_STEP_CM });
    expect(resolveTabAction(model, at(0, 0), true)).toEqual({ kind: "indent", deltaCm: -INDENT_STEP_CM });
  });
});

describe("resolveTabAction — rule 4, anywhere else", () => {
  it("types a literal tab mid-paragraph, even inside a list item (SPEC-054 AC-29/30)", () => {
    expect(resolveTabAction(doc(paragraph("text")), at(0, 2), false)).toEqual({ kind: "insertTab" });
    const listed = doc(paragraph("item", { type: "bullet", level: 1 }));
    expect(resolveTabAction(listed, at(0, 2), false)).toEqual({ kind: "insertTab" });
  });

  it("has nothing to insert for Shift+Tab mid-paragraph", () => {
    expect(resolveTabAction(doc(paragraph("text")), at(0, 2), true)).toEqual({ kind: "none" });
  });

  it("types a literal tab for a non-collapsed selection at offset 0", () => {
    const model = doc(paragraph("text"));
    const spanning: ModelSelection = { anchor: caret(0, 0), focus: caret(0, 4) };
    expect(resolveTabAction(model, spanning, false)).toEqual({ kind: "insertTab" });
  });
});
