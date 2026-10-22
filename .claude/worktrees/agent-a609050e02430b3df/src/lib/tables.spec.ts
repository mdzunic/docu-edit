import { describe, expect, it } from "vitest";
import {
  appendTableRowForTab,
  arrowEscapeCaret,
  caretTableLocation,
  deleteTable,
  deleteTableColumn,
  deleteTableRow,
  equalColumnWidthsCm,
  insertTable,
  insertTableColumn,
  insertTableRow,
  isInTableCell,
  nextCellCaret,
  prevCellCaret,
  tableAt,
} from "./tables";
import { cellRangeBounds, isCellInRange } from "./core/tables";
import { caretSelection } from "./commands";
import { equalStructural } from "./equality";
import { blockText, type CaretPosition } from "./textEditing";
import {
  createDocument,
  isTable,
  type DocumentModel,
  type Paragraph,
  type Table,
} from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function doc(...blocks: Array<Paragraph | Table>): DocumentModel {
  const model = createDocument("tables-fixture");
  model.sections = [{ blocks }];
  return model;
}

function cell(text: string) {
  return { blocks: [paragraph(text)] };
}

function grid(rows: string[][], widths?: number[]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({ cells: cells.map(cell) })),
    columnWidthsCm: widths ?? rows[0].map(() => 5),
  };
}

function caret(block: number, offset: number, cellPath?: { row: number; col: number; block: number }): CaretPosition {
  const position: CaretPosition = { section: 0, block, offset };
  if (cellPath) {
    position.cell = cellPath;
  }
  return position;
}

describe("equalColumnWidthsCm (AC-3)", () => {
  it("splits the text area (page minus margins) evenly, snapped to twips", () => {
    const model = createDocument("widths");
    const widths = equalColumnWidthsCm(model, 4);
    expect(widths).toHaveLength(4);
    // A4 (21 cm) minus 2 × 2.54 cm margins = 15.92 cm -> 3.98 cm each.
    for (const width of widths) {
      expect(width).toBeCloseTo(3.98, 2);
    }
    expect(new Set(widths).size).toBe(1);
  });
});

describe("insertTable (AC-3, AC-4)", () => {
  it("inserts before an empty caret paragraph, which becomes the following paragraph", () => {
    const model = doc(paragraph(""));
    const { model: next, selection } = insertTable(model, caretSelection(caret(0, 0)), 2, 3);
    expect(next.sections[0].blocks).toHaveLength(2);
    const table = next.sections[0].blocks[0];
    expect(isTable(table)).toBe(true);
    if (!isTable(table)) {
      return;
    }
    expect(table.rows).toHaveLength(2);
    expect(table.rows.every((row) => row.cells.length === 3)).toBe(true);
    expect(table.columnWidthsCm).toHaveLength(3);
    expect(next.sections[0].blocks[1].type).toBe("paragraph");
    // The caret lands in the first cell (AC-4).
    expect(selection.focus).toEqual(caret(0, 0, { row: 0, col: 0, block: 0 }));
  });

  it("splits a paragraph when inserting mid-text", () => {
    const model = doc(paragraph("hello world"));
    const { model: next, selection } = insertTable(model, caretSelection(caret(0, 5)), 1, 2);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(blocks[0])).toBe("hello");
    expect(blocks[1].type).toBe("table");
    expect(blockText(blocks[2])).toBe(" world");
    expect(selection.focus).toEqual(caret(1, 0, { row: 0, col: 0, block: 0 }));
  });

  it("keeps a paragraph after the table when inserting at a paragraph end (AC-14)", () => {
    const model = doc(paragraph("above"));
    const { model: next } = insertTable(model, caretSelection(caret(0, 5)), 1, 1);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blockText(blocks[0])).toBe("above");
    expect(blocks[1].type).toBe("table");
    expect(blocks[2].type).toBe("paragraph");
  });

  it("is a no-op for a caret inside a table (nested tables are out of scope)", () => {
    const model = doc(grid([["a"]]), paragraph(""));
    const inCell = caret(0, 0, { row: 0, col: 0, block: 0 });
    const { model: next } = insertTable(model, caretSelection(inCell), 2, 2);
    expect(next.sections[0].blocks).toHaveLength(2);
    expect(next.sections[0].blocks[0]).toBe(model.sections[0].blocks[0]);
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph("keep"));
    const snapshot = structuredClone(model);
    insertTable(model, caretSelection(caret(0, 2)), 2, 2);
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

describe("row and column commands (AC-17..AC-21)", () => {
  const location = { section: 0, block: 0, row: 0, col: 0 };

  it("insertTableRow above shifts the caret's row with it", () => {
    const model = doc(grid([["a", "b"], ["c", "d"]]), paragraph(""));
    const at = caret(0, 0, { row: 0, col: 1, block: 0 });
    const { model: next, selection } = insertTableRow(model, caretSelection(at), location, "above");
    const table = tableAt(next, 0, 0)!;
    expect(table.rows).toHaveLength(3);
    expect(table.rows[0].cells.map((c) => blockText(c.blocks[0]))).toEqual(["", ""]);
    expect(selection.focus.cell).toEqual({ row: 1, col: 1, block: 0 });
  });

  it("insertTableRow below leaves the caret in place", () => {
    const model = doc(grid([["a", "b"]]), paragraph(""));
    const at = caret(0, 0, { row: 0, col: 0, block: 0 });
    const { model: next, selection } = insertTableRow(model, caretSelection(at), location, "below");
    const table = tableAt(next, 0, 0)!;
    expect(table.rows).toHaveLength(2);
    expect(table.rows[1].cells.map((c) => blockText(c.blocks[0]))).toEqual(["", ""]);
    expect(selection.focus.cell).toEqual({ row: 0, col: 0, block: 0 });
  });

  it("insertTableColumn keeps the table's overall width", () => {
    const model = doc(grid([["a", "b"]], [6, 6]), paragraph(""));
    const { model: next } = insertTableColumn(model, caretSelection(caret(1, 0)), location, "right");
    const table = tableAt(next, 0, 0)!;
    expect(table.rows[0].cells).toHaveLength(3);
    expect(table.columnWidthsCm).toHaveLength(3);
    const total = table.columnWidthsCm.reduce((sum, cm) => sum + cm, 0);
    expect(total).toBeCloseTo(12, 1);
  });

  it("deleteTableRow moves the caret to the row that takes the deleted row's place", () => {
    const model = doc(grid([["a"], ["b"], ["c"]]), paragraph(""));
    const { model: next, selection } = deleteTableRow(
      model,
      caretSelection(caret(0, 0, { row: 1, col: 0, block: 0 })),
      { ...location, row: 1 },
    );
    const table = tableAt(next, 0, 0)!;
    expect(table.rows.map((row) => blockText(row.cells[0].blocks[0]))).toEqual(["a", "c"]);
    expect(selection.focus.cell).toEqual({ row: 1, col: 0, block: 0 });
  });

  it("deleting the last row deletes the table and lands the caret where it was (AC-22, AC-23)", () => {
    const model = doc(paragraph("before"), grid([["only"]]), paragraph("after"));
    const { model: next, selection } = deleteTableRow(
      model,
      caretSelection(caret(1, 0, { row: 0, col: 0, block: 0 })),
      { section: 0, block: 1, row: 0, col: 0 },
    );
    expect(next.sections[0].blocks.map(blockText)).toEqual(["before", "after"]);
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("deleting the last column deletes the table too", () => {
    const model = doc(grid([["x"], ["y"]]), paragraph("tail"));
    const { model: next, selection } = deleteTableColumn(model, caretSelection(caret(1, 0)), location);
    expect(next.sections[0].blocks.map(blockText)).toEqual(["tail"]);
    expect(selection.focus).toEqual(caret(0, 0));
  });

  it("deleteTableColumn removes the column and its width", () => {
    const model = doc(grid([["a", "b", "c"]], [3, 4, 5]), paragraph(""));
    const { model: next } = deleteTableColumn(model, caretSelection(caret(1, 0)), {
      ...location,
      col: 1,
    });
    const table = tableAt(next, 0, 0)!;
    expect(table.rows[0].cells.map((c) => blockText(c.blocks[0]))).toEqual(["a", "c"]);
    expect(table.columnWidthsCm).toEqual([3, 5]);
  });

  it("deleteTable removes the block; an emptied section keeps one paragraph", () => {
    const model = doc(grid([["solo"]]));
    const { model: next, selection } = deleteTable(model, caretSelection(caret(0, 0)), location);
    expect(next.sections[0].blocks).toHaveLength(1);
    expect(blockText(next.sections[0].blocks[0])).toBe("");
    expect(selection.focus).toEqual(caret(0, 0));
  });

  it("appendTableRowForTab adds a row and puts the caret in its first cell (AC-12)", () => {
    const model = doc(grid([["a", "b"]]), paragraph(""));
    const { model: next, selection } = appendTableRowForTab(model, caretSelection(caret(0, 0)), location);
    const table = tableAt(next, 0, 0)!;
    expect(table.rows).toHaveLength(2);
    expect(selection.focus.cell).toEqual({ row: 1, col: 0, block: 0 });
  });

  it("commands never mutate the input model", () => {
    const model = doc(grid([["a", "b"], ["c", "d"]]), paragraph(""));
    const snapshot = structuredClone(model);
    const selection = caretSelection(caret(0, 0, { row: 0, col: 0, block: 0 }));
    insertTableRow(model, selection, location, "below");
    insertTableColumn(model, selection, location, "right");
    deleteTableRow(model, selection, location);
    deleteTableColumn(model, selection, location);
    deleteTable(model, selection, location);
    appendTableRowForTab(model, selection, location);
    expect(equalStructural(model, snapshot)).toBe(true);
  });
});

describe("Tab navigation queries (AC-10, AC-11)", () => {
  const model = doc(paragraph("before"), grid([["a", "b"], ["c", "d"]]), paragraph("after"));

  it("nextCellCaret walks the grid in reading order and returns null after the last cell", () => {
    expect(nextCellCaret(model, caret(1, 0, { row: 0, col: 0, block: 0 }))?.cell).toEqual({
      row: 0,
      col: 1,
      block: 0,
    });
    expect(nextCellCaret(model, caret(1, 0, { row: 0, col: 1, block: 0 }))?.cell).toEqual({
      row: 1,
      col: 0,
      block: 0,
    });
    expect(nextCellCaret(model, caret(1, 0, { row: 1, col: 1, block: 0 }))).toBeNull();
  });

  it("prevCellCaret lands at the end of the previous cell and returns null in the first", () => {
    const prev = prevCellCaret(model, caret(1, 0, { row: 1, col: 0, block: 0 }));
    expect(prev?.cell).toEqual({ row: 0, col: 1, block: 0 });
    expect(prev?.offset).toBe(1);
    expect(prevCellCaret(model, caret(1, 0, { row: 0, col: 0, block: 0 }))).toBeNull();
  });

  it("caretTableLocation and isInTableCell recognise cell carets", () => {
    const inCell = caret(1, 0, { row: 1, col: 1, block: 0 });
    expect(isInTableCell(inCell)).toBe(true);
    expect(isInTableCell(caret(0, 0))).toBe(false);
    expect(caretTableLocation(model, inCell)).toEqual({ section: 0, block: 1, row: 1, col: 1 });
    expect(caretTableLocation(model, caret(0, 0))).toBeNull();
  });
});

describe("arrowEscapeCaret (AC-13, AC-14)", () => {
  const model = doc(paragraph("before"), grid([["a", "b"], ["c", "d"]]), paragraph("after"));

  it("down moves to the cell below, and out of the table from the last row", () => {
    expect(arrowEscapeCaret(model, caret(1, 0, { row: 0, col: 1, block: 0 }), "down")?.cell).toEqual({
      row: 1,
      col: 1,
      block: 0,
    });
    expect(arrowEscapeCaret(model, caret(1, 0, { row: 1, col: 0, block: 0 }), "down")).toEqual(
      caret(2, 0),
    );
  });

  it("up moves to the cell above, and to the block before the table from the first row", () => {
    expect(arrowEscapeCaret(model, caret(1, 0, { row: 1, col: 0, block: 0 }), "up")?.cell).toEqual({
      row: 0,
      col: 0,
      block: 0,
    });
    expect(arrowEscapeCaret(model, caret(1, 0, { row: 0, col: 0, block: 0 }), "up")).toEqual(
      caret(0, 6),
    );
  });

  it("right and left step through cells and exit at the table's ends", () => {
    expect(arrowEscapeCaret(model, caret(1, 1, { row: 0, col: 0, block: 0 }), "right")?.cell).toEqual({
      row: 0,
      col: 1,
      block: 0,
    });
    expect(arrowEscapeCaret(model, caret(1, 1, { row: 1, col: 1, block: 0 }), "right")).toEqual(
      caret(2, 0),
    );
    expect(arrowEscapeCaret(model, caret(1, 0, { row: 0, col: 0, block: 0 }), "left")).toEqual(
      caret(0, 6),
    );
  });

  it("returns null when there is nowhere to go (table starts the document)", () => {
    const edge = doc(grid([["x"]]), paragraph("after"));
    expect(arrowEscapeCaret(edge, caret(0, 0, { row: 0, col: 0, block: 0 }), "up")).toBeNull();
    expect(arrowEscapeCaret(edge, caret(0, 0, { row: 0, col: 0, block: 0 }), "left")).toBeNull();
  });
});

describe("cell ranges (AC-15, AC-16)", () => {
  it("normalises any drag direction into rectangular bounds", () => {
    const range = {
      section: 0,
      block: 1,
      anchor: { row: 2, col: 0 },
      focus: { row: 0, col: 3 },
    };
    expect(cellRangeBounds(range)).toEqual({ top: 0, left: 0, bottom: 2, right: 3 });
  });

  it("isCellInRange matches only cells of the range's own table", () => {
    const range = {
      section: 0,
      block: 1,
      anchor: { row: 0, col: 0 },
      focus: { row: 1, col: 1 },
    };
    expect(isCellInRange(range, 0, 1, 0, 1)).toBe(true);
    expect(isCellInRange(range, 0, 1, 1, 0)).toBe(true);
    expect(isCellInRange(range, 0, 1, 2, 0)).toBe(false);
    expect(isCellInRange(range, 0, 2, 0, 0)).toBe(false);
    expect(isCellInRange(null, 0, 1, 0, 0)).toBe(false);
  });
});
