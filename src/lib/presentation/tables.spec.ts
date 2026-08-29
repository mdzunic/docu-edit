import { describe, expect, it } from "vitest";
import { computeOwnerGrid, type CellRange } from "../core/tables";
import type { SlideTableCell, TableFrame } from "./model";
import {
  appendTableRow,
  applyCellBorderPreset,
  createTableFrame,
  deleteTableColumn,
  deleteTableRow,
  fullTableRange,
  insertTableColumn,
  insertTableRow,
  mergeCells,
  nextCellAddress,
  prevCellAddress,
  scaleTableGrid,
  setCellFill,
  setCellPadding,
  setCellVerticalAlign,
  setColumnWidths,
  setRowHeight,
  setTableBandRows,
  setTableHeaderRow,
  splitCell,
} from "./tables";

function makeTable(rows: number, cols: number): TableFrame {
  return createTableFrame(rows, cols, 20, 10);
}

function range(anchor: [number, number], focus: [number, number]): CellRange {
  return { section: 0, block: 0, anchor: { row: anchor[0], col: anchor[1] }, focus: { row: focus[0], col: focus[1] } };
}

/** Asserts every row is exactly columnWidthsCm.length wide and the owner grid
 * is regular (AC-9c). */
function expectRegularGrid(frame: TableFrame): void {
  const cols = frame.columnWidthsCm.length;
  for (const row of frame.rows) {
    expect(row.cells).toHaveLength(cols);
  }
  // computeOwnerGrid throwing/undefined would signal an irregular grid.
  const grid = computeOwnerGrid(frame);
  expect(grid).toHaveLength(frame.rows.length);
}

describe("createTableFrame (AC-5c)", () => {
  it("lands centred at 80% of slide width with equal columns and 0.9cm rows", () => {
    const t = createTableFrame(3, 4, 20, 10);
    expect(t.kind).toBe("table");
    expect(t.widthCm).toBeCloseTo(16, 3); // 80% of 20
    expect(t.xCm).toBeCloseTo(2, 3); // centred
    expect(t.columnWidthsCm).toHaveLength(4);
    const sum = t.columnWidthsCm.reduce((a, b) => a + b, 0);
    expect(sum).toBeCloseTo(16, 2);
    expect(new Set(t.columnWidthsCm).size).toBe(1); // equal
    expect(t.rows).toHaveLength(3);
    expect(t.rows.every((r) => r.heightCm === 0.9)).toBe(true);
    expect(t.heightCm).toBeCloseTo(2.7, 3);
    expect(t.yCm).toBeCloseTo((10 - 2.7) / 2, 3);
    expect(t.firstRow).toBe(true);
    expect(t.id).toBeTruthy();
  });

  it("clamps rows/cols to the grid limits", () => {
    const t = createTableFrame(9999, 9999, 20, 10);
    expect(t.rows.length).toBe(200);
    expect(t.columnWidthsCm.length).toBe(63);
  });
});

describe("row/column structure (AC-7)", () => {
  it("inserts a row above and below", () => {
    const t = makeTable(2, 2);
    insertTableRow(t, 0, "below");
    expect(t.rows).toHaveLength(3);
    insertTableRow(t, 0, "above");
    expect(t.rows).toHaveLength(4);
    expect(t.heightCm).toBeCloseTo(4 * 0.9, 3);
    expectRegularGrid(t);
  });

  it("appends a row for Tab in the last cell", () => {
    const t = makeTable(2, 3);
    appendTableRow(t);
    expect(t.rows).toHaveLength(3);
  });

  it("inserts a column taking the neighbour's width and grows widthCm to the sum (AC-7b)", () => {
    const t = makeTable(2, 2);
    t.columnWidthsCm = [4, 6];
    t.widthCm = 10;
    insertTableColumn(t, 1, "right");
    expect(t.columnWidthsCm).toEqual([4, 6, 6]); // new column takes neighbour's width
    expect(t.widthCm).toBeCloseTo(16, 3); // grown to the new sum
    expect(t.rows.every((r) => r.cells.length === 3)).toBe(true);
    expectRegularGrid(t);
  });

  it("deletes a row; deleting the last row signals frame removal (AC-7c)", () => {
    const t = makeTable(2, 2);
    expect(deleteTableRow(t, 0)).toBe(true);
    expect(t.rows).toHaveLength(1);
    expect(deleteTableRow(t, 0)).toBe(false); // last row → caller deletes frame
  });

  it("deletes a column and shrinks widthCm; last column signals removal", () => {
    const t = makeTable(2, 3);
    t.columnWidthsCm = [3, 4, 5];
    t.widthCm = 12;
    expect(deleteTableColumn(t, 1)).toBe(true);
    expect(t.columnWidthsCm).toEqual([3, 5]);
    expect(t.widthCm).toBeCloseTo(8, 3);
    expect(deleteTableColumn(t, 0)).toBe(true);
    expect(deleteTableColumn(t, 0)).toBe(false);
  });
});

describe("resize (AC-8)", () => {
  it("clamps dragged column widths to the minimum and sets widthCm to the sum", () => {
    const t = makeTable(1, 2);
    setColumnWidths(t, [0.1, 5]); // 0.1 is below MIN_COLUMN_WIDTH_CM
    expect(t.columnWidthsCm[0]).toBeCloseTo(0.5, 3);
    expect(t.widthCm).toBeCloseTo(5.5, 3);
  });

  it("sets a row height as a floor (min 0.4cm)", () => {
    const t = makeTable(2, 2);
    setRowHeight(t, 0, 0.1);
    expect(t.rows[0].heightCm).toBeCloseTo(0.4, 3);
    setRowHeight(t, 0, 3);
    expect(t.rows[0].heightCm).toBeCloseTo(3, 3);
  });

  it("scales the grid proportionally on a frame resize (AC-8c)", () => {
    const t = makeTable(2, 2);
    t.columnWidthsCm = [4, 4];
    scaleTableGrid(t, 2, 0.5);
    expect(t.columnWidthsCm).toEqual([8, 8]);
    expect(t.rows.every((r) => Math.abs((r.heightCm ?? 0) - 0.45) < 1e-6)).toBe(true);
  });
});

describe("merge / split (AC-9)", () => {
  it("merges a rectangular range into its top-left cell, concatenating content", () => {
    const t = makeTable(2, 2);
    t.rows[0].cells[0].paragraphs = [{ runs: [{ text: "A" }] }];
    t.rows[0].cells[1].paragraphs = [{ runs: [{ text: "B" }] }];
    t.rows[1].cells[0].paragraphs = [{ runs: [{ text: "C" }] }];
    t.rows[1].cells[1].paragraphs = [{ runs: [{ text: "D" }] }];
    mergeCells(t, range([0, 0], [1, 1]));
    const owner = t.rows[0].cells[0];
    expect(owner.colSpan).toBe(2);
    expect(owner.rowSpan).toBe(2);
    const text = owner.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("");
    expect(text).toBe("ABCD");
    expect(t.rows[0].cells[1].covered).toBe(true);
    expect(t.rows[1].cells[0].covered).toBe(true);
    expect(t.rows[1].cells[1].covered).toBe(true);
    expectRegularGrid(t);
  });

  it("addresses a merged cell from any slot via computeOwnerGrid (AC-9c)", () => {
    const t = makeTable(2, 2);
    mergeCells(t, range([0, 0], [1, 1]));
    const grid = computeOwnerGrid(t);
    expect(grid[1][1].originRow).toBe(0);
    expect(grid[1][1].originCol).toBe(0);
  });

  it("splits a merged cell back into a regular grid", () => {
    const t = makeTable(2, 2);
    mergeCells(t, range([0, 0], [1, 1]));
    splitCell(t, 0, 0, 2, 2);
    expect(t.rows[0].cells[0].colSpan).toBeUndefined();
    expect(t.rows[0].cells[0].rowSpan).toBeUndefined();
    expect(t.rows.every((r) => r.cells.every((c) => !c.covered))).toBe(true);
    expectRegularGrid(t);
  });

  it("grows the grid when splitting into more columns than the cell spans", () => {
    const t = makeTable(1, 1);
    t.rows[0].cells[0].paragraphs = [{ runs: [{ text: "X" }] }];
    splitCell(t, 0, 0, 1, 3);
    expect(t.columnWidthsCm).toHaveLength(3);
    expect(t.rows[0].cells).toHaveLength(3);
    expect(t.rows[0].cells[0].paragraphs[0].runs[0].text).toBe("X");
    expectRegularGrid(t);
  });
});

describe("per-cell formatting (AC-10) and preservation clearing (AC-12)", () => {
  it("sets and clears a cell fill, clearing preserved markup", () => {
    const t = makeTable(1, 2);
    (t.rows[0].cells[0] as SlideTableCell).pptxTcPr = "<a:tcPr/>";
    (t.rows[0].cells[0] as SlideTableCell).odpCellStyleName = "ce1";
    setCellFill(t, range([0, 0], [0, 0]), { type: "solid", color: "#ff0000" });
    expect(t.rows[0].cells[0].fill).toEqual({ type: "solid", color: "#ff0000" });
    expect(t.rows[0].cells[0].pptxTcPr).toBeUndefined();
    expect(t.rows[0].cells[0].odpCellStyleName).toBeUndefined();
    setCellFill(t, range([0, 0], [0, 0]), undefined);
    expect(t.rows[0].cells[0].fill).toBeUndefined();
  });

  it("applies an 'all' border preset to every edge of every cell in range", () => {
    const t = makeTable(2, 2);
    const side = { style: "single" as const, widthPt: 1, color: "#000000" };
    applyCellBorderPreset(t, fullTableRange(t), "all", side);
    for (const row of t.rows) {
      for (const cell of row.cells) {
        expect(cell.borders?.top).toEqual(side);
        expect(cell.borders?.bottom).toEqual(side);
      }
    }
  });

  it("sets padding and vertical alignment", () => {
    const t = makeTable(1, 1);
    setCellPadding(t, range([0, 0], [0, 0]), 0.25);
    expect(t.rows[0].cells[0].paddingCm).toBeCloseTo(0.25, 3);
    setCellVerticalAlign(t, range([0, 0], [0, 0]), "center");
    expect(t.rows[0].cells[0].verticalAlign).toBe("center");
  });

  it("toggles the header and banding look, clearing pptxTblPr", () => {
    const t = makeTable(2, 2);
    t.pptxTblPr = "<a:tblPr firstRow=\"1\"/>";
    setTableBandRows(t, true);
    expect(t.bandRow).toBe(true);
    expect(t.pptxTblPr).toBeUndefined();
    setTableHeaderRow(t, false);
    expect(t.firstRow).toBeUndefined();
  });
});

describe("cell navigation (AC-6b)", () => {
  it("moves to the next non-covered cell in reading order", () => {
    const t = makeTable(2, 2);
    expect(nextCellAddress(t, 0, 0)).toEqual({ row: 0, col: 1 });
    expect(nextCellAddress(t, 0, 1)).toEqual({ row: 1, col: 0 });
    expect(nextCellAddress(t, 1, 1)).toBeNull(); // last cell → append a row
  });

  it("skips covered slots after a merge", () => {
    const t = makeTable(2, 2);
    mergeCells(t, range([0, 0], [0, 1])); // merge the two cells of row 0
    // From the merged owner (0,0), next is row 1 col 0 (covered 0,1 skipped).
    expect(nextCellAddress(t, 0, 0)).toEqual({ row: 1, col: 0 });
    expect(prevCellAddress(t, 1, 0)).toEqual({ row: 0, col: 0 });
  });

  it("returns null before the first cell", () => {
    const t = makeTable(2, 2);
    expect(prevCellAddress(t, 0, 0)).toBeNull();
  });
});
