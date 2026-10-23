// SPEC-137 — the anchor index (anchors.ts) is carried from a sheet's previous
// rows array to its next when none of the rows that differ between the two
// holds an array's or a data table's anchor, so an edit or a recalculation
// slice on a large sheet re-reads the rows it touched instead of every cell;
// any differing row that held or holds an anchor rebuilds it. merges.ts's
// buildMergeIndex is counted to tell a carried index from a rebuilt one.
import { describe, expect, it, vi } from "vitest";
import type { Cell, Row, Sheet } from "../model";
import { findArrayAnchor, formulaAnchorAt } from "./anchors";

const built = vi.hoisted(() => ({ count: 0 }));

vi.mock("../merges", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../merges")>();
  return {
    ...actual,
    buildMergeIndex: (...args: Parameters<typeof actual.buildMergeIndex>) => {
      built.count += 1;
      return actual.buildMergeIndex(...args);
    },
  };
});

function sheetOf(id: string, rows: Row[]): Sheet {
  return { conditionalFormats: [], validations: [], id, name: id, visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

/** Rows 0..n-1 of plain numbers, an array anchored at B3 over B3:C5. */
function rowsWithArray(n: number): { rows: Row[]; anchor: Cell } {
  const anchor: Cell = { col: 1, value: 6, formula: "A3:A5*2", arrayRange: { top: 2, left: 1, bottom: 4, right: 2 } };
  const rows: Row[] = Array.from({ length: n }, (_, index) => ({ index, cells: [{ col: 0, value: index }] }));
  rows[2] = { index: 2, cells: [{ col: 0, value: 2 }, anchor] };
  return { rows, anchor };
}

/** `rows` with row `at`'s value replaced (copy-on-write: a new array and row). */
function edited(rows: Row[], at: number, cells: Cell[]): Row[] {
  const next = [...rows];
  next[at] = { index: rows[at]!.index, cells };
  return next;
}

describe("the anchor index across edits (SPEC-137)", () => {
  it("carries the index to the next rows array when no row that differs holds an anchor", () => {
    const { rows, anchor } = rowsWithArray(50);
    const before = sheetOf("carry", rows);
    expect(findArrayAnchor(before, 3, 2)).toBe(anchor);
    const builds = built.count;
    const after = sheetOf("carry", edited(rows, 40, [{ col: 0, value: 99 }]));
    expect(findArrayAnchor(after, 3, 2)).toBe(anchor);
    expect(findArrayAnchor(after, 40, 0)).toBeUndefined();
    // A recalculation slice's rewrite of many rows without an anchor carries it too.
    const slice = after.rows.map((row, index) => (index >= 10 && index < 30 ? { index: row.index, cells: [{ col: 0, value: -index }] } : row));
    expect(formulaAnchorAt(sheetOf("carry", slice), 4, 1)?.cell).toBe(anchor);
    expect(built.count).toBe(builds);
  });

  it("rebuilds when a row that differs held an anchor, or holds one now", () => {
    const { rows, anchor } = rowsWithArray(20);
    const before = sheetOf("rebuild", rows);
    expect(findArrayAnchor(before, 4, 2)).toBe(anchor);
    // The anchor's row rewritten without it: the array is gone.
    const cleared = sheetOf("rebuild", edited(rows, 2, [{ col: 0, value: 2 }, { col: 1, value: 6 }]));
    expect(findArrayAnchor(cleared, 4, 2)).toBeUndefined();
    // A new array anchored in another row: found.
    const moved: Cell = { col: 3, value: 1, formula: "1", arrayRange: { top: 10, left: 3, bottom: 11, right: 3 } };
    const added = sheetOf("rebuild", edited(cleared.rows, 10, [{ col: 0, value: 10 }, moved]));
    expect(findArrayAnchor(added, 11, 3)).toBe(moved);
    expect(findArrayAnchor(added, 4, 2)).toBeUndefined();
    // A data table's anchor joins the same way.
    const table: Cell = { col: 5, value: 0, dataTableXml: '<f t="dataTable" ref="F13:F14" dt2D="0" dtr="0" r1="A1"/>' };
    const withTable = sheetOf("rebuild", edited(added.rows, 12, [{ col: 0, value: 12 }, table]));
    expect(formulaAnchorAt(withTable, 13, 5)).toEqual({ kind: "dataTable", range: { top: 12, left: 5, bottom: 13, right: 5 }, cell: table });
    expect(findArrayAnchor(withTable, 11, 3)).toBe(moved);
  });

  it("rebuilds when rows shift past an anchor (an insertion renumbers the rows after it)", () => {
    const { rows, anchor } = rowsWithArray(10);
    expect(findArrayAnchor(sheetOf("shift", rows), 2, 1)).toBe(anchor);
    const shiftedAnchor: Cell = { ...anchor, arrayRange: { top: 3, left: 1, bottom: 5, right: 2 } };
    const shifted: Row[] = [rows[0]!, { index: 1, cells: [] }, ...rows.slice(1).map((row) => ({ index: row.index + 1, cells: row.index === 2 ? [{ col: 0, value: 2 }, shiftedAnchor] : row.cells }))];
    const after = sheetOf("shift", shifted);
    expect(findArrayAnchor(after, 2, 1)).toBeUndefined();
    expect(findArrayAnchor(after, 5, 2)).toBe(shiftedAnchor);
  });
});
