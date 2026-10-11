// SPEC-124 "Geometry of an edit", "The adjuster registry": every row of the
// geometry table, every structural command, the refusals, the inheritance of
// inserted rows and columns, copy-on-write, and the registry contract every
// later adjuster joins.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import { MAX_COLS, MAX_ROWS, type Cell, type ColumnSpan, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { MAX_FORMAT_MATERIALISED } from "./restyle";
import { autofitColumnWidthPx, fixedMeasurer } from "./autofit";
import { autofitColumns, autofitRows, freezePanes, MAX_ROW_HEIGHT_PT, resizeColumns, resizeRows, setColumnsHidden, setRowsHidden, TOO_MANY_ROWS_MESSAGE } from "./bands";
import type { ConditionalFormat } from "./conditionalFormats";
import { applyPattern, capturePattern, clearRange } from "./formatting";
import { applyPatternMerges, mergeCells, mergeDiscardsValues, patternMerges, unmergeCells, type MergeMode } from "./mergeCells";
import { rangesIntersect } from "./merges";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import type { AutoFilter } from "./filter";
import type { SortState } from "./sort";
import {
  adjustIndex,
  adjustRange,
  applyStructuralEdit,
  INSERT_PUSHES_DATA_MESSAGE,
  PARTIAL_ARRAY_MESSAGE,
  REFERENCE_ADJUSTERS,
  type StructuralEdit,
} from "./structure";

const LAST_ROW = MAX_ROWS - 1;
const LAST_COL = MAX_COLS - 1;

const rowsEdit = (kind: "insertRows" | "deleteRows", at: number, count: number, sheetId = "s1"): StructuralEdit => ({ kind, sheetId, at, count });
const colsEdit = (kind: "insertCols" | "deleteCols", at: number, count: number, sheetId = "s1"): StructuralEdit => ({ kind, sheetId, at, count });

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return {
    conditionalFormats: [],
    id,
    name,
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    ...fields,
  };
}

function model(sheets: Sheet[]): SpreadsheetModel {
  return {
    dxfs: [],
    kind: "spreadsheet",
    id: "m",
    title: "T",
    format: "xlsx",
    readerVersion: 8,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
  };
}

const value = (col: number, v: Cell["value"], extra: Partial<Cell> = {}): Cell => ({ col, value: v, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const sheetOf = (m: SpreadsheetModel, id = "s1"): Sheet => m.sheets.find((each) => each.id === id)!;
const indices = (s: Sheet): number[] => s.rows.map((r) => r.index);
const contains = (outer: GridRange, inner: GridRange): boolean =>
  outer.top <= inner.top && outer.left <= inner.left && outer.bottom >= inner.bottom && outer.right >= inner.right;

/** A small deterministic generator (mulberry32). */
function random(seed: number): () => number {
  let state = seed;
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let t = Math.imul(state ^ (state >>> 15), 1 | state);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe("adjustIndex — the geometry table's index rows", () => {
  it("insert n at a: i < a stays", () => {
    expect(adjustIndex(3, rowsEdit("insertRows", 5, 2))).toBe(3);
  });
  it("insert n at a: i ≥ a → i+n", () => {
    expect(adjustIndex(5, rowsEdit("insertRows", 5, 2))).toBe(7);
    expect(adjustIndex(9, colsEdit("insertCols", 5, 2))).toBe(11);
  });
  it("insert n at a: i+n past last → null", () => {
    expect(adjustIndex(LAST_ROW, rowsEdit("insertRows", 0, 1))).toBeNull();
    expect(adjustIndex(LAST_COL - 1, colsEdit("insertCols", 3, 2))).toBeNull();
    expect(adjustIndex(LAST_COL - 2, colsEdit("insertCols", 3, 2))).toBe(LAST_COL);
  });
  it("delete [a, a+n): i < a stays", () => {
    expect(adjustIndex(3, rowsEdit("deleteRows", 5, 2))).toBe(3);
  });
  it("delete [a, a+n): i in the band → null", () => {
    expect(adjustIndex(5, rowsEdit("deleteRows", 5, 2))).toBeNull();
    expect(adjustIndex(6, rowsEdit("deleteRows", 5, 2))).toBeNull();
  });
  it("delete [a, a+n): i ≥ a+n → i−n", () => {
    expect(adjustIndex(7, rowsEdit("deleteRows", 5, 2))).toBe(5);
  });
});

describe("adjustRange — the geometry table's range row", () => {
  const r = range(4, 1, 9, 3); // B5:D10
  it("insert: s ≥ a → s+n and e ≥ a → e+n (an insert at the first row moves the range)", () => {
    expect(adjustRange(r, rowsEdit("insertRows", 4, 2))).toEqual(range(6, 1, 11, 3));
  });
  it("insert strictly inside expands it", () => {
    expect(adjustRange(r, rowsEdit("insertRows", 6, 2))).toEqual(range(4, 1, 11, 3));
  });
  it("insert just after its last row leaves it alone — the same object", () => {
    expect(adjustRange(r, rowsEdit("insertRows", 10, 2))).toBe(r);
  });
  it("insert: the end clamps to last; a start pushed past last → null", () => {
    expect(adjustRange(range(0, 0, LAST_ROW, 0), rowsEdit("insertRows", 0, 1))).toEqual(range(1, 0, LAST_ROW, 0));
    expect(adjustRange(range(LAST_ROW - 1, 0, LAST_ROW, 0), rowsEdit("insertRows", 0, 2))).toBeNull();
  });
  it("delete: both ends in the band → null", () => {
    expect(adjustRange(range(4, 1, 5, 3), rowsEdit("deleteRows", 3, 4))).toBeNull();
  });
  it("delete: s in the band → a", () => {
    expect(adjustRange(r, rowsEdit("deleteRows", 3, 3))).toEqual(range(3, 1, 6, 3));
  });
  it("delete: e in the band → a−1", () => {
    expect(adjustRange(r, rowsEdit("deleteRows", 8, 4))).toEqual(range(4, 1, 7, 3));
  });
  it("delete: e ≥ a+n → e−n (a band strictly inside shrinks it)", () => {
    expect(adjustRange(r, rowsEdit("deleteRows", 6, 2))).toEqual(range(4, 1, 7, 3));
  });
  it("never touches the other axis", () => {
    expect(adjustRange(r, colsEdit("insertCols", 0, 5))).toEqual(range(4, 6, 9, 8));
    expect(adjustRange(r, colsEdit("deleteCols", 2, 1))).toEqual(range(4, 1, 9, 2));
  });
});

describe("applyStructuralEdit — inserting rows", () => {
  it("shifts every row at or after `at`, keeping the rows above by identity", () => {
    const above = row(1, [value(0, "a")]);
    const m = model([sheet("s1", "Sheet1", { rows: [above, row(2, [value(0, "b")]), row(5, [value(0, "c")])] })]);
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 2, 2))).toEqual({ ok: true });
    const s = sheetOf(m);
    expect(indices(s)).toEqual([1, 4, 7]);
    expect(s.rows[0]).toBe(above);
    expect(s.dirty).toBe(true);
  });

  it("new rows inherit the row above's styleId, heightPt, customHeight and non-zero cell styles as styled empty cells", () => {
    const source = row(1, [value(0, "x", { styleId: 3 }), value(1, 5), value(2, null, { styleId: 4 })], { styleId: 2, heightPt: 30, customHeight: true, hidden: true });
    const m = model([sheet("s1", "Sheet1", { rows: [source] })]);
    applyStructuralEdit(m, rowsEdit("insertRows", 2, 2));
    const inserted = sheetOf(m).rows.slice(1);
    expect(inserted).toEqual([
      { index: 2, styleId: 2, heightPt: 30, customHeight: true, cells: [{ col: 0, value: null, styleId: 3 }, { col: 2, value: null, styleId: 4 }] },
      { index: 3, styleId: 2, heightPt: 30, customHeight: true, cells: [{ col: 0, value: null, styleId: 3 }, { col: 2, value: null, styleId: 4 }] },
    ]);
  });

  it("at index 0 the new rows take the old first row's formats", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "Header", { styleId: 7 })], { heightPt: 22 })] })]);
    applyStructuralEdit(m, rowsEdit("insertRows", 0, 1));
    expect(sheetOf(m).rows).toEqual([
      { index: 0, heightPt: 22, cells: [{ col: 0, value: null, styleId: 7 }] },
      { index: 1, heightPt: 22, cells: [{ col: 0, value: "Header", styleId: 7 }] },
    ]);
  });

  it("copies only the row properties when the styled cells would pass MAX_FORMAT_MATERIALISED", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "x", { styleId: 3 }), value(1, "y", { styleId: 3 })], { heightPt: 20 })] })]);
    const count = Math.floor(MAX_FORMAT_MATERIALISED / 2) + 1;
    applyStructuralEdit(m, rowsEdit("insertRows", 1, count));
    const s = sheetOf(m);
    expect(s.rows[1]).toEqual({ index: 1, heightPt: 20, cells: [] });
    expect(s.rows).toHaveLength(count + 1);
  });

  it("stores nothing for an inserted row whose source stores nothing", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(4, [value(0, 1)])] })]);
    applyStructuralEdit(m, rowsEdit("insertRows", 2, 3));
    expect(indices(sheetOf(m))).toEqual([7]);
  });

  it("refuses an insert that would push a value or formula off the end", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(LAST_ROW, [value(0, "end")])] })]);
    const before = m.sheets;
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 3, 1))).toEqual({ ok: false, message: INSERT_PUSHES_DATA_MESSAGE });
    expect(m.sheets).toBe(before);
    const f = model([sheet("s1", "Sheet1", { rows: [row(LAST_ROW - 1, [value(0, null, { formula: "1+1" })])] })]);
    expect(applyStructuralEdit(f, rowsEdit("insertRows", 3, 2))).toEqual({ ok: false, message: INSERT_PUSHES_DATA_MESSAGE });
  });

  it("drops formatting-only rows and merges that fall off the end", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(2, [value(0, 1)]), row(LAST_ROW, [value(0, null, { styleId: 2 })], { heightPt: 30 })],
        merges: [range(LAST_ROW - 1, 0, LAST_ROW, 1), range(LAST_ROW, 2, LAST_ROW, 4)],
      }),
    ]);
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 0, 1)).ok).toBe(true);
    const s = sheetOf(m);
    expect(indices(s)).toEqual([3]);
    // The first merge clamps to one row of two cells; the second is gone.
    expect(s.merges).toEqual([range(LAST_ROW, 0, LAST_ROW, 1)]);
  });
});

describe("applyStructuralEdit — deleting rows", () => {
  it("removes the band's rows and shifts the rest up", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "a")]), row(3, [value(0, "b")]), row(4, [value(0, "c")]), row(9, [value(0, "d")])] })]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 3, 2));
    expect(sheetOf(m).rows).toEqual([row(0, [value(0, "a")]), row(7, [value(0, "d")])]);
  });

  it("removes merges inside the band, shrinks merges straddling it, and drops one reduced to a single cell", () => {
    const m = model([sheet("s1", "Sheet1", { merges: [range(3, 0, 4, 1), range(1, 2, 6, 2), range(4, 3, 5, 3), range(8, 0, 9, 1)] })]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 3, 2));
    expect(sheetOf(m).merges).toEqual([range(1, 2, 4, 2), range(6, 0, 7, 1)]);
  });

  it("deleting every row of a sheet leaves an empty sheet", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, 1)]), row(LAST_ROW, [value(0, 2)])], merges: [range(0, 0, 1, 1)] })]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 0, MAX_ROWS));
    const s = sheetOf(m);
    expect(s.rows).toEqual([]);
    expect(s.merges).toEqual([]);
  });

  it("turns every reference into the band #REF!, on every sheet", () => {
    const m = model([
      sheet("s1", "Sheet1", { rows: [row(0, [value(0, 3, { formula: "A5+A9" })])] }),
      sheet("s2", "Sheet2", { rows: [row(0, [value(0, 1, { formula: "Sheet1!A5" }), value(1, 2, { formula: "A5" })])] }),
    ]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 4, 1));
    expect(sheetOf(m).rows[0]!.cells[0]!.formula).toBe("#REF!+A8");
    expect(sheetOf(m, "s2").rows[0]!.cells.map((c) => c.formula)).toEqual(["Sheet1!#REF!", "A5"]);
    expect(sheetOf(m, "s2").dirty).toBe(true);
  });
});

describe("applyStructuralEdit — columns", () => {
  it("inserting shifts every cell at or after `at`; new columns take the left column's non-zero cell styles", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "a"), value(1, "b", { styleId: 5 }), value(2, "c")]), row(1, [value(2, "z")])] })]);
    applyStructuralEdit(m, colsEdit("insertCols", 2, 2));
    const [first, second] = sheetOf(m).rows;
    expect(first!.cells).toEqual([value(0, "a"), value(1, "b", { styleId: 5 }), value(2, null, { styleId: 5 }), value(3, null, { styleId: 5 }), value(4, "c")]);
    expect(second!.cells).toEqual([value(4, "z")]);
  });

  it("at index 0 the new columns take the old first column's style", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "a", { styleId: 2 })])] })]);
    applyStructuralEdit(m, colsEdit("insertCols", 0, 1));
    expect(sheetOf(m).rows[0]!.cells).toEqual([value(0, null, { styleId: 2 }), value(1, "a", { styleId: 2 })]);
  });

  it("refuses an insert pushing a filled cell past XFD", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(LAST_COL, 1)])] })]);
    expect(applyStructuralEdit(m, colsEdit("insertCols", 0, 1))).toEqual({ ok: false, message: INSERT_PUSHES_DATA_MESSAGE });
  });

  it("deleting removes the band's cells and a row left empty", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "a"), value(1, "b"), value(4, "e")]), row(1, [value(1, "only")])] })]);
    applyStructuralEdit(m, colsEdit("deleteCols", 1, 2));
    expect(sheetOf(m).rows).toEqual([row(0, [value(0, "a"), value(2, "e")])]);
  });

  it("columns: spans shift, split for an insert inside, copy the left span's width and style but not hidden, and re-join", () => {
    const spans: ColumnSpan[] = [
      { min: 1, max: 4, widthPx: 100, styleId: 2 },
      { min: 6, max: 6, hidden: true, widthPx: 30 },
    ];
    const m = model([sheet("s1", "Sheet1", { columns: spans })]);
    applyStructuralEdit(m, colsEdit("insertCols", 3, 2));
    expect(sheetOf(m).columns).toEqual([
      { min: 1, max: 6, widthPx: 100, styleId: 2 },
      { min: 8, max: 8, hidden: true, widthPx: 30 },
    ]);
    const h = model([sheet("s1", "Sheet1", { columns: [{ min: 0, max: 2, hidden: true, widthPx: 30 }] })]);
    applyStructuralEdit(h, colsEdit("insertCols", 2, 1));
    expect(sheetOf(h).columns).toEqual([
      { min: 0, max: 1, hidden: true, widthPx: 30 },
      { min: 2, max: 2, widthPx: 30 },
      { min: 3, max: 3, hidden: true, widthPx: 30 },
    ]);
  });

  it("columns: a delete trims and removes spans and re-joins what meets", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        columns: [
          { min: 0, max: 1, widthPx: 80 },
          { min: 2, max: 3, widthPx: 120 },
          { min: 4, max: 6, widthPx: 80 },
        ],
      }),
    ]);
    applyStructuralEdit(m, colsEdit("deleteCols", 2, 2));
    expect(sheetOf(m).columns).toEqual([{ min: 0, max: 4, widthPx: 80 }]);
  });
});

describe("applyStructuralEdit — refusals", () => {
  it("refuses an edit splitting an array formula's range, before anything changes", () => {
    const anchor = value(1, 2, { formula: "A1:A2*2", arrayRange: range(4, 1, 6, 1) });
    const m = model([sheet("s1", "Sheet1", { rows: [row(4, [anchor])] })]);
    const before = m.sheets;
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 5, 1))).toEqual({ ok: false, message: PARTIAL_ARRAY_MESSAGE });
    expect(applyStructuralEdit(m, rowsEdit("deleteRows", 6, 2))).toEqual({ ok: false, message: PARTIAL_ARRAY_MESSAGE });
    expect(m.sheets).toBe(before);
    // At its first row it moves; past its last row, or around it whole, it is fine.
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 4, 1)).ok).toBe(true);
    expect(sheetOf(m).rows[0]!.cells[0]!.arrayRange).toEqual(range(5, 1, 7, 1));
    expect(applyStructuralEdit(m, rowsEdit("deleteRows", 3, 6)).ok).toBe(true);
    expect(sheetOf(m).rows).toEqual([]);
  });

  it("refuses on a protected sheet unless the protection permits the operation", () => {
    const m = model([sheet("s1", "Sheet1", { protection: { sheet: true }, rows: [row(3, [value(0, 1)])] })]);
    for (const edit of [rowsEdit("insertRows", 0, 1), rowsEdit("deleteRows", 0, 1), colsEdit("insertCols", 0, 1), colsEdit("deleteCols", 0, 1)]) {
      expect(applyStructuralEdit(m, edit)).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    }
    const permitted = model([sheet("s1", "Sheet1", { protection: { sheet: true, allowInsertRows: true, allowDeleteColumns: true } })]);
    expect(applyStructuralEdit(permitted, rowsEdit("insertRows", 0, 1)).ok).toBe(true);
    expect(applyStructuralEdit(permitted, colsEdit("deleteCols", 0, 1)).ok).toBe(true);
    expect(applyStructuralEdit(permitted, rowsEdit("deleteRows", 0, 1))).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    expect(applyStructuralEdit(permitted, colsEdit("insertCols", 0, 1))).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
  });
});

describe("applyStructuralEdit — the other adjusters", () => {
  it("moves arrays, frozen panes, the selection and the formulas of every sheet", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(9, [value(0, 1, { formula: "SUM(A1:A9)", arrayRange: range(9, 0, 10, 0) })])],
        frozen: { rows: 3, cols: 1 },
        selection: { active: { row: 5, col: 2 }, ranges: [range(5, 2, 6, 3)], activeRangeIndex: 0 },
      }),
    ]);
    applyStructuralEdit(m, rowsEdit("insertRows", 1, 2));
    const s = sheetOf(m);
    expect(s.rows[0]!.cells[0]).toEqual(value(0, 1, { formula: "SUM(A1:A11)", arrayRange: range(11, 0, 12, 0) }));
    expect(s.frozen).toEqual({ rows: 5, cols: 1 });
    expect(s.selection).toEqual({ active: { row: 7, col: 2 }, ranges: [range(7, 2, 8, 3)], activeRangeIndex: 0 });
  });

  it("unfreezes a band the delete removes, and keeps a selection whose rows are deleted where it was", () => {
    const m = model([
      sheet("s1", "Sheet1", { frozen: { rows: 2, cols: 0 }, selection: { active: { row: 0, col: 0 }, ranges: [range(0, 0, 1, LAST_COL)], activeRangeIndex: 0 } }),
    ]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 0, 2));
    const s = sheetOf(m);
    expect(s.frozen).toBeUndefined();
    expect(s.selection).toEqual({ active: { row: 0, col: 0 }, ranges: [range(0, 0, 1, LAST_COL)], activeRangeIndex: 0 });
  });

  it("never touches formulaUnparsed text, and leaves a sheet nothing touches by identity", () => {
    const wire = value(0, 1, { formula: "of:=[.A5]", formulaUnparsed: true });
    const other = sheet("s2", "Sheet2", { rows: [row(0, [value(0, 2, { formula: "B2" })])] });
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [wire])] }), other]);
    applyStructuralEdit(m, rowsEdit("insertRows", 1, 1));
    expect(sheetOf(m).rows[0]!.cells[0]).toBe(wire);
    expect(sheetOf(m, "s2")).toBe(other);
  });

  it("is copy-on-write: no array or object the model held before is mutated", () => {
    const rows = [row(0, [value(0, 1)]), row(2, [value(0, 2)])];
    const merges = [range(2, 0, 3, 1)];
    const columns: ColumnSpan[] = [{ min: 2, max: 2, widthPx: 90 }];
    const frozen = { rows: 3, cols: 3 };
    const s = sheet("s1", "Sheet1", { rows, merges, columns, frozen });
    const snapshot = structuredClone(s);
    const m = model([s]);
    applyStructuralEdit(m, rowsEdit("insertRows", 1, 1));
    applyStructuralEdit(m, colsEdit("deleteCols", 0, 1));
    expect(s).toEqual(snapshot);
    expect(sheetOf(m)).not.toBe(s);
  });
});

// --- the band commands (bands.ts) and merges (mergeCells.ts) ---------------------------------

const CTX = { locale: SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE], dateSystem: "1900" as const };

describe("resize, hide, unhide (SPEC-124 \"Sizes\")", () => {
  it("resizeRows writes heightPt rounded to 0.75 pt with customHeight, creating property-only rows", () => {
    const stored = row(1, [value(0, 1)], { sourceStyleName: "ro2" });
    const m = model([sheet("s1", "Sheet1", { rows: [stored] })]);
    expect(resizeRows(m, "s1", [{ start: 1, end: 2 }], 20.2)).toEqual({ ok: true });
    expect(sheetOf(m).rows).toEqual([row(1, [value(0, 1)], { heightPt: 20.25, customHeight: true }), row(2, [], { heightPt: 20.25, customHeight: true })]);
    expect(sheetOf(m).dirty).toBe(true);
    resizeRows(m, "s1", [{ start: 1, end: 1 }], 1000);
    expect(sheetOf(m).rows[0]!.heightPt).toBe(MAX_ROW_HEIGHT_PT);
  });

  it("a height of 0 (or under 4 px) hides the rows and keeps their size", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(3, [], { heightPt: 30, customHeight: true })] })]);
    resizeRows(m, "s1", [{ start: 3, end: 3 }], 0);
    expect(sheetOf(m).rows).toEqual([row(3, [], { heightPt: 30, customHeight: true, hidden: true })]);
    resizeRows(m, "s1", [{ start: 3, end: 3 }], 18);
    expect(sheetOf(m).rows).toEqual([row(3, [], { heightPt: 18, customHeight: true })]);
  });

  it("resizeColumns writes widthPx and drops sourceWidthChars unless the typed width is given; under 4 px hides", () => {
    const m = model([sheet("s1", "Sheet1", { columns: [{ min: 1, max: 3, widthPx: 90, sourceWidthChars: 12.7109375 }] })]);
    resizeColumns(m, "s1", [{ start: 2, end: 2 }], 120);
    expect(sheetOf(m).columns).toEqual([
      { min: 1, max: 1, widthPx: 90, sourceWidthChars: 12.7109375 },
      { min: 2, max: 2, widthPx: 120 },
      { min: 3, max: 3, widthPx: 90, sourceWidthChars: 12.7109375 },
    ]);
    resizeColumns(m, "s1", [{ start: 5, end: 5 }], 78, 11.2109375);
    expect(sheetOf(m).columns.at(-1)).toEqual({ min: 5, max: 5, widthPx: 78, sourceWidthChars: 11.2109375 });
    resizeColumns(m, "s1", [{ start: 2, end: 2 }], 3);
    expect(sheetOf(m).columns[1]).toEqual({ min: 2, max: 2, widthPx: 120, hidden: true });
  });

  it("SPEC-128: Unhide clears `hidden` only — a filtered-out row stays hidden until the filter changes (a sheet holding both kinds)", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(2, [value(0, "West")], { filterHidden: true }), row(3, [value(0, "East")], { hidden: true }), row(4, [], { hidden: true, filterHidden: true })],
      }),
    ]);
    expect(setRowsHidden(m, "s1", [{ start: 0, end: 10 }], false)).toEqual({ ok: true });
    expect(sheetOf(m).rows).toEqual([row(2, [value(0, "West")], { filterHidden: true }), row(3, [value(0, "East")]), row(4, [], { filterHidden: true })]);
    // Hiding a filtered-out row adds the user's flag beside the filter's.
    setRowsHidden(m, "s1", [{ start: 2, end: 2 }], true);
    expect(sheetOf(m).rows[0]).toEqual(row(2, [value(0, "West")], { filterHidden: true, hidden: true }));
  });

  it("hide and unhide touch only `hidden`; unhiding stores nothing new", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(4, [], { heightPt: 40, customHeight: true })], columns: [{ min: 2, max: 2, widthPx: 100 }] })]);
    setRowsHidden(m, "s1", [{ start: 4, end: 5 }], true);
    expect(sheetOf(m).rows).toEqual([row(4, [], { heightPt: 40, customHeight: true, hidden: true }), row(5, [], { hidden: true })]);
    setRowsHidden(m, "s1", [{ start: 0, end: 10 }], false);
    expect(sheetOf(m).rows).toEqual([row(4, [], { heightPt: 40, customHeight: true })]);
    setColumnsHidden(m, "s1", [{ start: 1, end: 2 }], true);
    expect(sheetOf(m).columns).toEqual([{ min: 1, max: 1, hidden: true }, { min: 2, max: 2, widthPx: 100, hidden: true }]);
    setColumnsHidden(m, "s1", [{ start: 1, end: 2 }], false);
    expect(sheetOf(m).columns).toEqual([{ min: 2, max: 2, widthPx: 100 }]);
  });

  it("refuses hiding more than MAX_FORMAT_MATERIALISED rows that store nothing, and a protected sheet without the format permit", () => {
    const m = model([sheet("s1", "Sheet1")]);
    expect(setRowsHidden(m, "s1", [{ start: 0, end: MAX_FORMAT_MATERIALISED }], true)).toEqual({ ok: false, message: TOO_MANY_ROWS_MESSAGE });
    expect(setRowsHidden(m, "s1", [{ start: 0, end: MAX_FORMAT_MATERIALISED - 1 }], true).ok).toBe(true);
    const locked = model([sheet("s1", "Sheet1", { protection: { sheet: true, allowFormatColumns: true } })]);
    expect(resizeRows(locked, "s1", [{ start: 0, end: 0 }], 30)).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    expect(resizeColumns(locked, "s1", [{ start: 0, end: 0 }], 30).ok).toBe(true);
  });

  it("freezePanes sets and clears Sheet.frozen", () => {
    const m = model([sheet("s1", "Sheet1")]);
    freezePanes(m, "s1", 1, 0);
    expect(sheetOf(m).frozen).toEqual({ rows: 1, cols: 0 });
    freezePanes(m, "s1", 0, 0);
    expect(sheetOf(m).frozen).toBeUndefined();
  });
});

describe("autofit (SPEC-124 \"Autofit\") through fixedMeasurer", () => {
  const measurer = fixedMeasurer(7, 20);

  it("autofitColumnWidthPx: the widest display text + 5 px, skipping wrapped and merged cells; no text → the default", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(0, [value(0, "abcdefghij"), value(1, "x".repeat(40), { styleId: 1 }), value(2, "merged text here")]), row(1, [value(0, 12345)])],
        merges: [range(0, 2, 0, 3)],
      }),
    ]);
    m.styles = [{}, { alignment: { wrap: true } }];
    const s = sheetOf(m);
    expect(autofitColumnWidthPx(m, s, 0, measurer, CTX)).toBe(10 * 7 + 5);
    expect(autofitColumnWidthPx(m, s, 1, measurer, CTX)).toBe(64);
    expect(autofitColumnWidthPx(m, s, 2, measurer, CTX)).toBe(64);
    expect(autofitColumnWidthPx(m, s, 9, measurer, CTX)).toBe(64);
  });

  it("caps a column at the width of 255 characters", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "w".repeat(2000))])] })]);
    // 255 characters at the default font's MDW (7 px): trunc((255 + trunc(128/7)/256) × 7).
    expect(autofitColumnWidthPx(m, sheetOf(m), 0, measurer, CTX)).toBe(1785);
  });

  it("autofitColumns writes the measured widths; autofitRows the measured height with customHeight cleared", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "abcdefghijklmn")], { heightPt: 40, customHeight: true })] })]);
    autofitColumns(m, "s1", [{ start: 0, end: 1 }], measurer, CTX);
    expect(sheetOf(m).columns).toEqual([{ min: 0, max: 0, widthPx: 14 * 7 + 5 }]);
    autofitRows(m, "s1", [{ start: 0, end: 0 }], fixedMeasurer(7, 40), CTX);
    expect(sheetOf(m).rows[0]).toEqual(row(0, [value(0, "abcdefghijklmn")], { heightPt: 30 }));
    // A measured height at or under the default stores none.
    autofitRows(m, "s1", [{ start: 0, end: 0 }], fixedMeasurer(7, 10), CTX);
    expect(sheetOf(m).rows[0]).toEqual(row(0, [value(0, "abcdefghijklmn")]));
  });
});

describe("merges (SPEC-124 \"Merges\")", () => {
  const filled = (): SpreadsheetModel => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(0, [value(0, "Title"), value(1, "B1", { styleId: 1 }), value(2, null, { formula: "1+1" })]), row(1, [value(0, "a"), value(1, "b")])],
      }),
    ]);
    m.styles = [{}, { font: { bold: true } }];
    return m;
  };

  it("mergeCells 'cells' keeps the upper-left value, clears the others' values and formulas, keeps their styles", () => {
    const m = filled();
    expect(mergeDiscardsValues(m, "s1", [range(0, 0, 0, 2)], "cells")).toBe(true);
    expect(mergeCells(m, "s1", [range(0, 0, 0, 2)], "cells")).toEqual({ ok: true });
    const s = sheetOf(m);
    expect(s.merges).toEqual([range(0, 0, 0, 2)]);
    expect(s.rows[0]!.cells).toEqual([value(0, "Title"), value(1, null, { styleId: 1 })]);
    expect(s.dirty).toBe(true);
  });

  it("'centre' adds horizontal center alignment to the upper-left cell", () => {
    const m = filled();
    mergeCells(m, "s1", [range(0, 0, 0, 1)], "centre");
    const anchor = sheetOf(m).rows[0]!.cells[0]!;
    expect(m.styles[anchor.styleId!]!.alignment).toEqual({ horizontal: "center" });
  });

  it("'across' merges each row separately; a single cell is a no-op", () => {
    const m = filled();
    mergeCells(m, "s1", [range(0, 0, 1, 1)], "across");
    expect(sheetOf(m).merges).toEqual([range(0, 0, 0, 1), range(1, 0, 1, 1)]);
    const single = filled();
    const sheets = single.sheets;
    expect(mergeCells(single, "s1", [range(1, 1, 1, 1)], "cells")).toEqual({ ok: true });
    expect(single.sheets).toBe(sheets);
  });

  it("'across' over a merge spanning several rows replaces it with one merge per row", () => {
    // A1:B2 is merged: A1:C2 merged across leaves A1:C1 and A2:C2, never A1:B2 beside them.
    const m = model([sheet("s1", "Sheet1", { merges: [range(0, 0, 1, 1), range(5, 5, 6, 6)] })]);
    expect(mergeCells(m, "s1", [range(0, 0, 1, 2)], "across")).toEqual({ ok: true });
    expect(sheetOf(m).merges).toEqual([range(5, 5, 6, 6), range(0, 0, 0, 2), range(1, 0, 1, 2)]);
    // A cell of the merge alone expands to all of it, split by rows as well.
    const inside = model([sheet("s1", "Sheet1", { merges: [range(0, 0, 1, 2)] })]);
    expect(mergeCells(inside, "s1", [range(1, 1, 1, 1)], "across")).toEqual({ ok: true });
    expect(sheetOf(inside).merges).toEqual([range(0, 0, 0, 2), range(1, 0, 1, 2)]);
  });

  it("'across' over a merge one column wide removes it — each of its rows is a single cell", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "top")]), row(1, [value(0, "kept")])], merges: [range(0, 0, 2, 0)] })]);
    const rows = sheetOf(m).rows;
    expect(mergeDiscardsValues(m, "s1", [range(0, 0, 2, 0)], "across")).toBe(false);
    expect(mergeCells(m, "s1", [range(0, 0, 2, 0)], "across")).toEqual({ ok: true });
    expect(sheetOf(m).merges).toEqual([]);
    expect(sheetOf(m).rows).toBe(rows);
  });

  it("ranges are merged one after another: a range meeting an earlier one's merge extends over it", () => {
    const m = model([sheet("s1", "Sheet1")]);
    expect(mergeCells(m, "s1", [range(0, 0, 0, 1), range(0, 1, 1, 2)], "cells")).toEqual({ ok: true });
    expect(sheetOf(m).merges).toEqual([range(0, 0, 1, 2)]);
    const across = model([sheet("s1", "Sheet1")]);
    expect(mergeCells(across, "s1", [range(0, 0, 1, 2), range(1, 1, 2, 3)], "across")).toEqual({ ok: true });
    expect(sheetOf(across).merges).toEqual([range(0, 0, 0, 2), range(1, 0, 1, 3), range(2, 0, 2, 3)]);
  });

  it("refuses to unmerge a protected sheet's locked cells, as merging them is refused", () => {
    const m = model([sheet("s1", "Sheet1", { protection: { sheet: true }, merges: [range(0, 0, 2, 0)] })]);
    const sheets = m.sheets;
    expect(mergeCells(m, "s1", [range(0, 0, 2, 0)], "across")).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    expect(m.sheets).toBe(sheets);
  });

  it("a range partially overlapping a merge extends to cover it, and the merges inside are replaced", () => {
    const m = model([sheet("s1", "Sheet1", { merges: [range(1, 1, 2, 2), range(5, 5, 6, 6)] })]);
    expect(mergeCells(m, "s1", [range(0, 0, 1, 1)], "cells")).toEqual({ ok: true });
    expect(sheetOf(m).merges).toEqual([range(5, 5, 6, 6), range(0, 0, 2, 2)]);
  });

  it("refuses a merge partly over an array range", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(1, 2, { formula: "A1:A2*2", arrayRange: range(0, 1, 1, 1) })])] })]);
    expect(mergeCells(m, "s1", [range(0, 0, 0, 1)], "cells")).toEqual({ ok: false, message: PARTIAL_ARRAY_MESSAGE });
  });

  it("unmergeCells removes every intersecting merge, content and styles kept", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "x", { styleId: 0 })])], merges: [range(0, 0, 1, 1), range(4, 4, 5, 5)] })]);
    const rows = sheetOf(m).rows;
    unmergeCells(m, "s1", [range(1, 1, 1, 1)]);
    expect(sheetOf(m).merges).toEqual([range(4, 4, 5, 5)]);
    expect(sheetOf(m).rows).toBe(rows);
  });

  it("Clear formats and Clear all unmerge the merges inside the cleared ranges", () => {
    const m = model([sheet("s1", "Sheet1", { merges: [range(0, 0, 0, 1), range(3, 0, 3, 5)] })]);
    clearRange(m, "s1", [range(0, 0, 1, 2)], "formats");
    expect(sheetOf(m).merges).toEqual([range(3, 0, 3, 5)]);
    clearRange(m, "s1", [range(3, 0, 3, 5)], "all");
    expect(sheetOf(m).merges).toEqual([]);
  });

  it("the format painter reproduces the source's merges in the target", () => {
    const m = model([sheet("s1", "Sheet1", { merges: [range(0, 0, 0, 1)], rows: [row(0, [value(0, "x")]), row(5, [value(3, "keep"), value(4, "lost")])] })]);
    const pattern = capturePattern(sheetOf(m), range(0, 0, 0, 1))!;
    expect(pattern.merges).toEqual([range(0, 0, 0, 1)]);
    expect(patternMerges(range(5, 3, 5, 6), pattern)).toEqual([range(5, 3, 5, 4), range(5, 5, 5, 6)]);
    expect(mergeDiscardsValues(m, "s1", patternMerges(range(5, 3, 5, 6), pattern), "cells")).toBe(true);
    applyPattern(m, "s1", range(5, 3, 5, 6), pattern);
    applyPatternMerges(m, "s1", range(5, 3, 5, 6), pattern);
    expect(sheetOf(m).merges).toEqual([range(0, 0, 0, 1), range(5, 3, 5, 4), range(5, 5, 5, 6)]);
    expect(sheetOf(m).rows[1]!.cells).toEqual([value(3, "keep")]);
  });

  it("the painter removes a merge reaching into the target that a placed merge would overlap", () => {
    // A1:B1 painted on the single cell D6 covers D6:E6, which E6:F7 reaches into.
    const m = model([sheet("s1", "Sheet1", { merges: [range(0, 0, 0, 1), range(5, 4, 6, 5), range(9, 0, 9, 1)] })]);
    const pattern = capturePattern(sheetOf(m), range(0, 0, 0, 1))!;
    expect(applyPatternMerges(m, "s1", range(5, 3, 5, 3), pattern)).toEqual({ ok: true });
    expect(sheetOf(m).merges).toEqual([range(0, 0, 0, 1), range(9, 0, 9, 1), range(5, 3, 5, 4)]);
  });

  it("never leaves two merges overlapping, and every row of a selected range lies in a merge (randomised)", () => {
    const next = random(124);
    const int = (n: number): number => Math.floor(next() * n);
    const area = (): GridRange => {
      const top = int(10);
      const left = int(10);
      return range(top, left, top + int(4), left + int(4));
    };
    const modes: MergeMode[] = ["centre", "across", "cells"];
    for (let k = 0; k < 400; k += 1) {
      const existing: GridRange[] = [];
      for (let j = 0; j < 6; j += 1) {
        const merge = area();
        if ((merge.top !== merge.bottom || merge.left !== merge.right) && !existing.some((done) => rangesIntersect(done, merge))) {
          existing.push(merge);
        }
      }
      const m = model([sheet("s1", "Sheet1", { merges: existing })]);
      const ranges = Array.from({ length: 1 + int(3) }, area);
      const mode = modes[int(modes.length)]!;
      expect(mergeCells(m, "s1", ranges, mode)).toEqual({ ok: true });
      const merges = sheetOf(m).merges;
      const scenario = JSON.stringify({ existing, ranges, mode, merges });
      merges.forEach((a, i) => merges.slice(i + 1).forEach((b) => expect(rangesIntersect(a, b), scenario).toBe(false)));
      // "across" leaves a range one column wide unmerged, row by row.
      for (const selected of ranges.filter((each) => each.left !== each.right || (mode !== "across" && each.top !== each.bottom))) {
        const pieces = mode === "across" ? Array.from({ length: selected.bottom - selected.top + 1 }, (_, d) => ({ ...selected, top: selected.top + d, bottom: selected.top + d })) : [selected];
        for (const piece of pieces) {
          expect(merges.some((merge) => contains(merge, piece)), scenario).toBe(true);
        }
      }
    }
  });
});

// --- SPEC-128's autofilter adjuster ("Structural edits") -------------------------------------

const sortOf = (ref: GridRange, cols: number[], extra: Partial<SortState> = {}): SortState => ({ ref, keys: cols.map((col) => ({ col, descending: false })), ...extra });

describe("the autofilter adjuster (SPEC-128 \"Structural edits\")", () => {
  /** An autofilter at A5:D20 with criteria on B and D and a sort by D. */
  const filterAt = (fields: Partial<AutoFilter> = {}): AutoFilter => ({
    range: range(4, 0, 19, 3),
    columns: [
      { colId: 1, criteria: { kind: "values", values: ["West"], blanks: false, dateGroups: [] } },
      { colId: 3, criteria: { kind: "top10", top: true, percent: false, value: 3 } },
    ],
    sort: sortOf(range(5, 0, 19, 3), [3], { sourceXml: `<sortState ref="A6:D20"><sortCondition ref="D6:D20"/></sortState>` }),
    sourceXml: `<autoFilter ref="A5:D20"/>`,
    ...fields,
  });
  const filtered = (fields: Partial<Sheet> = {}): SpreadsheetModel =>
    model([sheet("s1", "Sheet1", { autoFilter: filterAt(), rows: [row(4, [value(0, "Region")]), row(6, [value(0, "West")], { filterHidden: true }), row(8, [value(0, "East")], { hidden: true })], ...fields })]);

  it("inserting two rows above an autofilter at A5:D20 moves it to A7:D22 with its sort record and the worksheet's sortState, without their source", () => {
    const m = filtered({ sortState: sortOf(range(4, 5, 9, 6), [6], { sourceXml: `<sortState ref="F5:G10"/>` }) });
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 1, 2))).toEqual({ ok: true });
    const s = sheetOf(m);
    expect(s.autoFilter).toEqual({ range: range(6, 0, 21, 3), columns: filterAt().columns, sort: sortOf(range(7, 0, 21, 3), [3]) });
    expect(s.sortState).toEqual(sortOf(range(6, 5, 11, 6), [6]));
    // The rows it hid moved with their rows; the user's hidden row too.
    expect(s.rows.map((r) => [r.index, r.filterHidden, r.hidden])).toEqual([
      [6, undefined, undefined],
      [8, true, undefined],
      [10, undefined, true],
    ]);
    expect(m.structureDirty).toBe(true);
  });

  it("leaves an autofilter and a sort record an edit does not reach by identity (and the workbook part alone)", () => {
    const m = filtered({ sortState: sortOf(range(4, 5, 9, 6), [6]) });
    const { autoFilter, sortState } = sheetOf(m);
    applyStructuralEdit(m, rowsEdit("insertRows", 30, 2));
    applyStructuralEdit(m, colsEdit("insertCols", 10, 1));
    expect(sheetOf(m).autoFilter).toBe(autoFilter);
    expect(sheetOf(m).sortState).toBe(sortState);
    expect(m.structureDirty).toBeUndefined();
  });

  it("a row inserted inside the data rows grows the range and is visible, though the row above it is filtered out", () => {
    const m = filtered();
    sheetOf(m).rows[1] = row(6, [value(0, "West", { styleId: 1 })], { filterHidden: true, heightPt: 30 });
    applyStructuralEdit(m, rowsEdit("insertRows", 7, 1));
    const s = sheetOf(m);
    expect(s.autoFilter!.range).toEqual(range(4, 0, 20, 3));
    expect(s.rows.find((r) => r.index === 7)).toEqual(row(7, [value(0, null, { styleId: 1 })], { heightPt: 30 }));
  });

  it("a column inserted inside the range shifts the colIds (and sort keys) at or right of it; one at its left edge moves the range only", () => {
    const m = filtered();
    const left = sheetOf(m).autoFilter!.columns[0];
    applyStructuralEdit(m, colsEdit("insertCols", 2, 1));
    let filter = sheetOf(m).autoFilter!;
    expect(filter.range).toEqual(range(4, 0, 19, 4));
    expect(filter.columns.map((column) => column.colId)).toEqual([1, 4]);
    expect(filter.columns[0]).toBe(left);
    expect(filter.sort).toEqual(sortOf(range(5, 0, 19, 4), [4]));
    expect(filter.sourceXml).toBeUndefined();
    applyStructuralEdit(m, colsEdit("insertCols", 0, 2));
    filter = sheetOf(m).autoFilter!;
    expect(filter.range).toEqual(range(4, 2, 19, 6));
    expect(filter.columns.map((column) => column.colId)).toEqual([1, 4]);
    expect(filter.sort!.keys).toEqual([{ col: 6, descending: false }]);
  });

  it("deleting a criterion's column drops that criterion and its sort keys; a sort record left with no keys is removed", () => {
    const m = filtered({ sortState: sortOf(range(30, 0, 40, 3), [3]) });
    applyStructuralEdit(m, colsEdit("deleteCols", 1, 1));
    let s = sheetOf(m);
    expect(s.autoFilter!.range).toEqual(range(4, 0, 19, 2));
    expect(s.autoFilter!.columns).toEqual([{ colId: 2, criteria: { kind: "top10", top: true, percent: false, value: 3 } }]);
    expect(s.autoFilter!.sort).toEqual(sortOf(range(5, 0, 19, 2), [2]));
    applyStructuralEdit(m, colsEdit("deleteCols", 2, 1));
    s = sheetOf(m);
    expect(s.autoFilter!.columns).toEqual([]);
    expect(s.autoFilter!.sort).toBeUndefined();
    expect("sort" in s.autoFilter!).toBe(false);
    expect(s.sortState).toBeUndefined();
    expect("sortState" in s).toBe(false);
  });

  it("deleting the header row removes the autofilter, and the rows it hid stay hidden as the user's", () => {
    const m = filtered();
    applyStructuralEdit(m, rowsEdit("deleteRows", 4, 1));
    const s = sheetOf(m);
    expect(s.autoFilter).toBeUndefined();
    expect("autoFilter" in s).toBe(false);
    expect(s.rows.map((r) => [r.index, r.filterHidden, r.hidden])).toEqual([
      [5, undefined, true],
      [7, undefined, true],
    ]);
    expect(m.structureDirty).toBe(true);
  });

  it("deleting data rows shrinks the range; deleting every column of it removes it", () => {
    const m = filtered();
    applyStructuralEdit(m, rowsEdit("deleteRows", 10, 20));
    expect(sheetOf(m).autoFilter!.range).toEqual(range(4, 0, 9, 3));
    applyStructuralEdit(m, colsEdit("deleteCols", 0, 4));
    expect(sheetOf(m).autoFilter).toBeUndefined();
  });

  it("a column edit reshaping a range with an ods filter that does not decompose drops it; moving the range keeps it", () => {
    const unmapped = `<table:filter><table:filter-or/></table:filter>`;
    const m = filtered({ autoFilter: filterAt({ odsName: "__Anonymous_Sheet_DB__0", odsUnmappedFilter: unmapped }) });
    applyStructuralEdit(m, colsEdit("insertCols", 0, 1));
    expect(sheetOf(m).autoFilter).toMatchObject({ odsName: "__Anonymous_Sheet_DB__0", odsUnmappedFilter: unmapped });
    applyStructuralEdit(m, colsEdit("insertCols", 2, 1));
    expect(sheetOf(m).autoFilter!.odsUnmappedFilter).toBeUndefined();
  });

  it("is copy-on-write: the autofilter, its columns and sort records are replaced, never mutated", () => {
    const m = filtered({ sortState: sortOf(range(4, 5, 9, 6), [6]) });
    const before = sheetOf(m);
    const snapshot = structuredClone(before);
    applyStructuralEdit(m, colsEdit("insertCols", 2, 1));
    applyStructuralEdit(m, rowsEdit("deleteRows", 4, 1));
    expect(before).toEqual(snapshot);
  });
});

// --- SPEC-129's conditionalFormats adjuster ("Structural edits") -----------------------------

/** A format over B5:B20 with one expression rule, self-referential to its
 * origin (B5) and an absolute cell ($Z$1) — SPEC-129's own example. */
const expressionFormat = (fields: Partial<ConditionalFormat> = {}): ConditionalFormat => ({
  id: "f1",
  ranges: [range(4, 1, 19, 1)],
  rules: [{ id: "r1", priority: 1, kind: "expression", formula: "B5>$Z$1" }],
  ...fields,
});

describe("the conditionalFormats adjuster (SPEC-129 \"Structural edits\")", () => {
  it("inserting two rows above a format on B5:B20 moves it to B7:B22 and its expression B5>$Z$1 to B7>$Z$1", () => {
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [expressionFormat()] })]);
    expect(applyStructuralEdit(m, rowsEdit("insertRows", 2, 2))).toEqual({ ok: true });
    const format = sheetOf(m).conditionalFormats[0]!;
    expect(format.ranges).toEqual([range(6, 1, 21, 1)]);
    expect(format.rules[0]).toMatchObject({ kind: "expression", formula: "B7>$Z$1" });
  });

  it("deleting the format's first row re-anchors the formula onto the surviving top-left (no #REF!)", () => {
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [expressionFormat()] })]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 4, 1));
    const format = sheetOf(m).conditionalFormats[0]!;
    expect(format.ranges).toEqual([range(4, 1, 18, 1)]); // B5:B19
    expect(format.rules[0]).toMatchObject({ kind: "expression", formula: "B5>$Z$1" });
  });

  it("drops a fully deleted range and keeps the other; removes an empty format and renumbers priorities", () => {
    const twoRanges = expressionFormat({ ranges: [range(4, 1, 19, 1), range(30, 1, 35, 1)] });
    const second = expressionFormat({ id: "f2", ranges: [range(40, 1, 40, 1)], rules: [{ id: "r2", priority: 2, kind: "expression", formula: "B41>0" }] });
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [twoRanges, second] })]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 30, 6));
    let formats = sheetOf(m).conditionalFormats;
    expect(formats).toHaveLength(2);
    expect(formats[0]!.ranges).toEqual([range(4, 1, 19, 1)]); // the second range (30:35) is gone, the first stays
    expect(formats[0]!.rules[0]!.priority).toBe(1);
    expect(formats[1]!.rules[0]!.priority).toBe(2);

    // Deleting the whole first format's remaining range removes it and renumbers.
    applyStructuralEdit(m, rowsEdit("deleteRows", 4, 16));
    formats = sheetOf(m).conditionalFormats;
    expect(formats).toHaveLength(1);
    expect(formats[0]!.id).toBe("f2");
    expect(formats[0]!.rules[0]!.priority).toBe(1);
  });

  it("re-anchors onto the next standing range when the first is deleted whole, and follows the column axis", () => {
    // B5:B6 and D10:D20; deleting rows 5–6 leaves D10:D20 (now D8:D18) leading.
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [expressionFormat({ ranges: [range(4, 1, 5, 1), range(9, 3, 19, 3)] })] })]);
    applyStructuralEdit(m, rowsEdit("deleteRows", 4, 2));
    expect(sheetOf(m).conditionalFormats[0]!.ranges).toEqual([range(7, 3, 17, 3)]);
    expect(sheetOf(m).conditionalFormats[0]!.rules[0]).toMatchObject({ formula: "D8>$Z$1" });
    // A format on B5:D5: deleting column B re-anchors onto C5, now B5 ($Z$1, right of the band, moves left too).
    const cols = model([sheet("s1", "Sheet1", { conditionalFormats: [expressionFormat({ ranges: [range(4, 1, 4, 3)] })] })]);
    applyStructuralEdit(cols, { kind: "deleteCols", sheetId: "s1", at: 1, count: 1 });
    expect(sheetOf(cols).conditionalFormats[0]!.ranges).toEqual([range(4, 1, 4, 2)]);
    expect(sheetOf(cols).conditionalFormats[0]!.rules[0]).toMatchObject({ formula: "B5>$Y$1" });
  });

  it("adjusts a format on another sheet whose formula references the edited sheet, and marks it dirty", () => {
    const other: ConditionalFormat = { id: "f2", ranges: [range(0, 0, 0, 0)], rules: [{ id: "r2", priority: 1, kind: "expression", formula: "A1>Sheet1!$Z$1" }] };
    const m = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2", { conditionalFormats: [other] })]);
    applyStructuralEdit(m, rowsEdit("insertRows", 0, 3));
    const s2 = sheetOf(m, "s2");
    expect(s2.conditionalFormats[0]!.rules[0]).toMatchObject({ formula: "A1>Sheet1!$Z$4" });
    expect(s2.dirty).toBe(true);
  });

  it("adjusts a colour scale's value-object formula and a preserved rule's embedded <formula>, and leaves an untouched format by identity", () => {
    const colorScale: ConditionalFormat = {
      id: "cs",
      ranges: [range(59, 1, 69, 1)], // B60:B70
      rules: [{ id: "r", priority: 1, kind: "colorScale", stops: [{ at: { type: "num", value: "$Z$60" }, color: "#ff0000" }, { at: { type: "max" }, color: "#00ff00" }] }],
    };
    const preserved: ConditionalFormat = {
      id: "pv",
      ranges: [range(59, 2, 69, 2)], // C60:C70
      rules: [{ id: "rp", priority: 2, kind: "preserved", where: "base", xml: `<cfRule type="expression"><formula>$Z$60&gt;0</formula></cfRule>` }],
    };
    // A distinct priority: ranks must stay unique across the sheet's formats.
    const untouched = expressionFormat({ id: "far", rules: [{ id: "r1", priority: 3, kind: "expression", formula: "B5>$Z$1" }] }); // B5:B20, well above the inserted band
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [colorScale, preserved, untouched] })]);
    applyStructuralEdit(m, rowsEdit("insertRows", 50, 1));
    const formats = sheetOf(m).conditionalFormats;
    const byId = (id: string): ConditionalFormat => formats.find((format) => format.id === id)!;
    expect(byId("cs").ranges).toEqual([range(60, 1, 70, 1)]);
    expect(byId("cs").rules[0]).toMatchObject({ stops: [{ at: { type: "num", value: "$Z$61" } }, { at: { type: "max" } }] });
    expect(byId("pv").rules[0]).toMatchObject({ xml: `<cfRule type="expression"><formula>$Z$61&gt;0</formula></cfRule>` });
    expect(byId("far")).toBe(untouched);
  });

  it("is copy-on-write: a touched format loses its sourceXml, an untouched one keeps it, nothing is mutated", () => {
    const touched = expressionFormat({ sourceXml: "<conditionalFormatting/>" });
    const snapshot = structuredClone(touched);
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [touched] })]);
    applyStructuralEdit(m, rowsEdit("insertRows", 2, 1));
    expect(touched).toEqual(snapshot);
    expect(sheetOf(m).conditionalFormats[0]!.sourceXml).toBeUndefined();
  });

  it("leaves a model with no conditional formats untouched by identity", () => {
    const m = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")]);
    const sheets = m.sheets;
    applyStructuralEdit(m, rowsEdit("insertRows", 0, 1));
    expect(m.sheets[1]).toBe(sheets[1]);
  });
});

// --- the registry contract ------------------------------------------------------------------
// Every adjuster joins this table: a model holding no structure of its kind is
// left untouched by identity, and a row inserted above its structure moves it.
// A later spec appends its row with its adjuster.

/** A model holding one structure of each kind this spec moves, at row 5. */
function withStructure(id: string): SpreadsheetModel {
  const base = sheet("s1", "Sheet1");
  switch (id) {
    case "rows":
      return model([{ ...base, rows: [row(5, [value(0, 1)])] }]);
    case "formulas":
      return model([{ ...base, rows: [row(0, [value(0, 1, { formula: "A6" })])] }]);
    case "merges":
      return model([{ ...base, merges: [range(5, 0, 6, 1)] }]);
    case "frozen":
      return model([{ ...base, frozen: { rows: 6, cols: 0 } }]);
    case "selection":
      return model([{ ...base, selection: { active: { row: 5, col: 0 }, ranges: [range(5, 0, 5, 0)], activeRangeIndex: 0 } }]);
    case "columns":
      // Columns move only under column edits: the "row" inserted is column 0.
      return model([{ ...base, columns: [{ min: 5, max: 5, widthPx: 90 }] }]);
    case "retained":
      return model([base]);
    case "autofilter":
      return model([{ ...base, autoFilter: { range: range(5, 0, 7, 1), columns: [] } }]);
    case "conditionalFormats":
      return model([{ ...base, conditionalFormats: [{ id: "f1", ranges: [range(5, 0, 5, 0)], rules: [{ id: "r1", priority: 1, kind: "expression", formula: "A6>0" }] }] }]);
    default:
      throw new Error(`no contract fixture for the "${id}" adjuster`);
  }
}

/** Where the structure of each kind is, to see it move. */
function structureOf(id: string, m: SpreadsheetModel): unknown {
  const s = sheetOf(m);
  switch (id) {
    case "rows":
      return s.rows[0]!.index;
    case "formulas":
      return s.rows[0]!.cells[0]!.formula;
    case "merges":
      return s.merges[0];
    case "frozen":
      return s.frozen;
    case "selection":
      return s.selection?.active;
    case "columns":
      return s.columns[0];
    case "autofilter":
      return s.autoFilter?.range;
    case "conditionalFormats":
      return s.conditionalFormats[0]?.ranges[0];
    default:
      return undefined;
  }
}

describe("REFERENCE_ADJUSTERS — the registry contract", () => {
  it("registers this spec's adjusters in order, then SPEC-128's autofilter, then SPEC-129's conditionalFormats", () => {
    expect(REFERENCE_ADJUSTERS.map((adjuster) => adjuster.id)).toEqual(["rows", "formulas", "merges", "frozen", "selection", "columns", "retained", "autofilter", "conditionalFormats"]);
  });

  for (const adjuster of REFERENCE_ADJUSTERS) {
    it(`${adjuster.id}: no structure of its kind → the model untouched by identity`, () => {
      const m = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")]);
      const sheets = m.sheets;
      const [first, second] = sheets;
      const opaque = m.opaque;
      adjuster.adjust(m, rowsEdit("insertRows", 0, 1));
      adjuster.adjust(m, colsEdit("insertCols", 0, 1));
      expect(m.sheets).toBe(sheets);
      expect(m.sheets[0]).toBe(first);
      expect(m.sheets[1]).toBe(second);
      expect(m.opaque).toBe(opaque);
    });

    if (adjuster.id !== "retained") {
      // retained's own contract is its format halves' specs (xlsx/ and
      // ods/retainedReferences.spec.ts): a model with no retained source
      // has nothing to move.
      it(`${adjuster.id}: a row inserted above its structure moves it`, () => {
        const m = withStructure(adjuster.id);
        const before = structureOf(adjuster.id, m);
        const edit = adjuster.id === "columns" ? colsEdit("insertCols", 2, 1) : rowsEdit("insertRows", 2, 1);
        adjuster.adjust(m, edit);
        const after = structureOf(adjuster.id, m);
        expect(after).not.toEqual(before);
        expect(after).toEqual(
          {
            rows: 6,
            formulas: "A7",
            merges: range(6, 0, 7, 1),
            frozen: { rows: 7, cols: 0 },
            selection: { row: 6, col: 0 },
            columns: { min: 6, max: 6, widthPx: 90 },
            autofilter: range(6, 0, 8, 1),
            conditionalFormats: range(6, 0, 6, 0),
          }[adjuster.id],
        );
      });
    }
  }
});
