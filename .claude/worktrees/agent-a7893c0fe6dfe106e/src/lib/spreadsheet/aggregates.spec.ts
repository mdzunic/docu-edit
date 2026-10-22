// SPEC-117 "Aggregates": the status bar's Average/Count/Sum over a
// selection's stored cells — each cell once where ranges overlap, hidden rows
// and columns excluded, booleans and numeric-looking strings counted but
// never summed, an error suppressing Average and Sum — shown only for a
// selection of more than one cell with a value in it, numbers through
// generalText. Whole-column and whole-sheet selections cost the stored rows
// they cross (proven by counting row reads), never the grid's positions.
import { describe, expect, it } from "vitest";
import { parseRange, type GridRange } from "./address";
import {
  formatAggregates,
  selectionAggregates,
  selectionSpansMultipleCells,
  selectionStoredCellEstimate,
  type SelectionAggregates,
} from "./aggregates";
import { buildGridMetrics } from "./geometry";
import { buildMergeIndex } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type ColumnSpan, type Row, type Sheet } from "./model";
import type { NavContext } from "./selection";

function range(text: string): GridRange {
  const parsed = parseRange(text);
  if (!parsed) {
    throw new Error(`bad range ${text}`);
  }
  return parsed;
}

/** A sheet from `{ "B3": 10, … }`, with hidden rows (0-based) and columns. */
function sheetOf(cells: Record<string, CellValue>, hidden: { rows?: number[]; cols?: number[] } = {}): Sheet {
  const byRow = new Map<number, Cell[]>();
  for (const [a1, value] of Object.entries(cells)) {
    const at = range(a1);
    byRow.set(at.top, [...(byRow.get(at.top) ?? []), { col: at.left, value }]);
  }
  for (const index of hidden.rows ?? []) {
    byRow.set(index, byRow.get(index) ?? []);
  }
  const hiddenRows = new Set(hidden.rows ?? []);
  const rows = [...byRow.entries()]
    .toSorted(([a], [b]) => a - b)
    .map(([index, rowCells]): Row => ({ index, cells: rowCells.toSorted((a, b) => a.col - b.col), ...(hiddenRows.has(index) ? { hidden: true } : {}) }));
  const columns = (hidden.cols ?? []).toSorted((a, b) => a - b).map((col): ColumnSpan => ({ min: col, max: col, hidden: true }));
  return { conditionalFormats: [], validations: [], id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows, columns, defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

function visibility(sheet: Sheet): Pick<NavContext, "isRowHidden" | "isColHidden"> {
  const metrics = buildGridMetrics(sheet);
  return { isRowHidden: (r) => metrics.rowHeight(r) === 0, isColHidden: (c) => metrics.colWidth(c) === 0 };
}

function aggregate(sheet: Sheet, ...ranges: string[]): SelectionAggregates {
  return selectionAggregates(sheet, ranges.map(range), visibility(sheet));
}

const WHOLE_SHEET: GridRange = { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 };

describe("selectionAggregates", () => {
  it("10, 20, 30 → count 3, sum 60", () => {
    const sheet = sheetOf({ B3: 10, B4: 20, B5: 30, B6: 40 });
    expect(aggregate(sheet, "B3:B5")).toEqual({ count: 3, numericCount: 3, sum: 60, hasError: false });
  });

  it("counts each cell once where ranges overlap", () => {
    const sheet = sheetOf({ A1: 1, B1: 1, C1: 1, A2: 1, B2: 1, C2: 1, A3: 1, B3: 1, C3: 1 });
    expect(aggregate(sheet, "A1:B2", "B2:C3")).toMatchObject({ count: 7, sum: 7 });
    expect(aggregate(sheet, "A1:C3", "A1:C3", "B2")).toMatchObject({ count: 9, sum: 9 });
    expect(aggregate(sheet, "A1:A3", "C1:C3", "A2:C2")).toMatchObject({ count: 7, sum: 7 });
  });

  it("excludes hidden rows and columns", () => {
    const sheet = sheetOf({ A1: 1, B1: 2, C1: 4, A2: 8, B2: 16, C2: 32 }, { rows: [1], cols: [1] });
    expect(aggregate(sheet, "A1:C2")).toEqual({ count: 2, numericCount: 2, sum: 5, hasError: false });
  });

  it("counts booleans, strings and numeric-looking strings but sums numbers only", () => {
    const sheet = sheetOf({ A1: 5, A2: true, A3: "12", A4: "text", A5: "", A6: 0.5 });
    expect(aggregate(sheet, "A1:A6")).toEqual({ count: 6, numericCount: 2, sum: 5.5, hasError: false });
  });

  it("skips style-only cells (value null) and flags errors", () => {
    const sheet = sheetOf({ A1: 1, A2: null, A3: { error: "#DIV/0!" } });
    expect(aggregate(sheet, "A1:A3")).toEqual({ count: 2, numericCount: 1, sum: 1, hasError: true });
    expect(aggregate(sheet, "A2")).toEqual({ count: 0, numericCount: 0, sum: 0, hasError: false });
  });

  it("handles whole-column and whole-sheet ranges over far-apart stored rows", () => {
    const sheet = sheetOf({ A1: 1, B1: 2, A500000: 3, XFD1048576: 4 });
    expect(aggregate(sheet, "A1:A1048576")).toMatchObject({ count: 2, sum: 4 });
    expect(selectionAggregates(sheet, [WHOLE_SHEET], visibility(sheet))).toMatchObject({ count: 4, sum: 10 });
    expect(aggregate(sheet, "C2:XFD1048575")).toMatchObject({ count: 0 });
  });

  it("normalises and clamps ranges, dropping ones outside the grid", () => {
    const sheet = sheetOf({ A1: 1, B2: 2 });
    const odd: GridRange[] = [
      { top: 1, left: 1, bottom: 0, right: 0 },
      { top: -5, left: -5, bottom: 0, right: 0 },
      { top: MAX_ROWS + 5, left: 0, bottom: MAX_ROWS + 9, right: 3 },
    ];
    expect(selectionAggregates(sheet, odd, visibility(sheet))).toMatchObject({ count: 2, sum: 3 });
  });

  it("reaches a band by binary search: a few whole rows of a 1,000-row sheet read few rows", () => {
    const rows: Row[] = Array.from({ length: 1000 }, (_, index) => ({ index, cells: [{ col: 0, value: 1 }] }));
    let reads = 0;
    const counted = new Proxy(rows, {
      get(target, property, receiver) {
        if (typeof property === "string" && /^\d+$/.test(property)) {
          reads += 1;
        }
        return Reflect.get(target, property, receiver);
      },
    });
    const sheet: Sheet = { ...sheetOf({}), rows: counted };
    const totals = selectionAggregates(sheet, [{ top: 500, left: 0, bottom: 509, right: MAX_COLS - 1 }], visibility(sheetOf({})));
    expect(totals.count).toBe(10);
    expect(reads).toBeLessThan(60);
  });
});

describe("selectionStoredCellEstimate", () => {
  it("is an upper bound of the stored cells visited, tight for rows inside the range", () => {
    const sheet = sheetOf({ A1: 1, B1: 2, C1: 3, A2: 4, A3: 5, Z3: 6 });
    expect(selectionStoredCellEstimate(sheet, [range("A1:C3")])).toBeGreaterThanOrEqual(aggregate(sheet, "A1:C3").count);
    expect(selectionStoredCellEstimate(sheet, [WHOLE_SHEET])).toBe(6);
    expect(selectionStoredCellEstimate(sheet, [range("A1:A3")])).toBe(3);
    expect(selectionStoredCellEstimate(sheet, [range("A1:C3"), range("A1:C3")])).toBe(selectionStoredCellEstimate(sheet, [range("A1:C3")]));
    expect(selectionStoredCellEstimate(sheet, [range("A10:C20")])).toBe(0);
  });

  it("stops counting once past the cap", () => {
    const rows: Row[] = Array.from({ length: 10_000 }, (_, index) => ({ index, cells: [{ col: 0, value: index }] }));
    const sheet: Sheet = { ...sheetOf({}), rows };
    expect(selectionStoredCellEstimate(sheet, [WHOLE_SHEET])).toBe(10_000);
    expect(selectionStoredCellEstimate(sheet, [WHOLE_SHEET], 100)).toBe(101);
  });
});

describe("selectionSpansMultipleCells", () => {
  it("is false for one cell (or one merged cell) and true for anything larger", () => {
    expect(selectionSpansMultipleCells([])).toBe(false);
    expect(selectionSpansMultipleCells([range("B3")])).toBe(false);
    expect(selectionSpansMultipleCells([range("B3:B4")])).toBe(true);
    expect(selectionSpansMultipleCells([range("B3"), range("D5")])).toBe(true);
    const merges = buildMergeIndex([range("C5:E6")]);
    expect(selectionSpansMultipleCells([range("C5:E6")], merges)).toBe(false);
    expect(selectionSpansMultipleCells([range("C5:E6")])).toBe(true);
    expect(selectionSpansMultipleCells([range("B5:E6")], merges)).toBe(true);
  });
});

describe("formatAggregates", () => {
  it("10, 20, 30 → \"Average: 20\", \"Count: 3\", \"Sum: 60\"", () => {
    const sheet = sheetOf({ B3: 10, B4: 20, B5: 30 });
    expect(formatAggregates(aggregate(sheet, "B3:B5"), true)).toEqual({ average: "Average: 20", count: "Count: 3", sum: "Sum: 60" });
  });

  it("shows nothing for a single cell or a selection without values", () => {
    expect(formatAggregates({ count: 1, numericCount: 1, sum: 5, hasError: false }, false)).toEqual({});
    expect(formatAggregates({ count: 0, numericCount: 0, sum: 0, hasError: false }, true)).toEqual({});
  });

  it("shows only the count without numbers or with an error", () => {
    expect(formatAggregates({ count: 2, numericCount: 0, sum: 0, hasError: false }, true)).toEqual({ count: "Count: 2" });
    expect(formatAggregates({ count: 3, numericCount: 2, sum: 3, hasError: true }, true)).toEqual({ count: "Count: 3" });
  });

  it("formats numbers through generalText", () => {
    expect(formatAggregates({ count: 3, numericCount: 3, sum: 1, hasError: false }, true)).toEqual({
      average: "Average: 0.333333333",
      count: "Count: 3",
      sum: "Sum: 1",
    });
    const tenths = sheetOf({ A1: 0.1, A2: 0.2 });
    expect(formatAggregates(aggregate(tenths, "A1:A2"), true)).toMatchObject({ average: "Average: 0.15", sum: "Sum: 0.3" });
    expect(formatAggregates({ count: 2, numericCount: 2, sum: 123_456_789_012, hasError: false }, true)).toMatchObject({ sum: "Sum: 1.23457E+11" });
  });
});
