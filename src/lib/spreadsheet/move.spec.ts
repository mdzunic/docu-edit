// SPEC-125 "Cut and move": every row of the re-point table, #REF! for
// references to overwritten cells, moved formulas not shifted, overlapping
// source and destination, cross-sheet moves, the refusals, moveMapper for
// retained text, and the registry contract every later MoveAdjuster joins.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { getCell } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import { PASTE_MERGE_MESSAGE, PASTE_PAST_EDGE_MESSAGE, TOO_MANY_TO_CHANGE_MESSAGE } from "./clipboard";
import type { AutoFilter } from "./filter";
import { MAX_ROWS, type Cell, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { MOVE_ADJUSTERS, moveMapper, moveRange, moveRefusal, repointFormula, snappedMove, type RangeMove } from "./move";
import { PARTIAL_ARRAY_MESSAGE } from "./structure";

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], id, name, visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(sheets: Sheet[], fields: Partial<SpreadsheetModel> = {}): SpreadsheetModel {
  return { dxfs: [], kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 8, sheets, activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0, ...fields };
}

const value = (col: number, v: Cell["value"], extra: Partial<Cell> = {}): Cell => ({ col, value: v, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const sheetOf = (m: SpreadsheetModel, id: string): Sheet => m.sheets.find((each) => each.id === id)!;
const cellOf = (m: SpreadsheetModel, id: string, r: number, c: number): Cell | undefined => getCell(sheetOf(m, id), r, c);

/** A10:A12 → H10:H12 on Sheet1 (s1). */
const DOWN: RangeMove = { fromSheetId: "s1", from: range(9, 0, 11, 0), toSheetId: "s1", toTop: 9, toLeft: 7 };
const names: Record<string, string> = { s1: "Sheet1", s2: "Sheet2", s3: "Other Sheet" };
const nameOf = (id: string): string => names[id] ?? "";

describe("repointFormula — the re-point table", () => {
  const repoint = (text: string, host = "s2", move: RangeMove = DOWN): string => repointFormula(text, host, move, nameOf);

  it("moves a reference entirely inside the source by the offset, absolute or relative", () => {
    expect(repoint("SUM(Sheet1!A10:A12)")).toBe("SUM(Sheet1!H10:H12)");
    expect(repoint("Sheet1!$A$10+Sheet1!A11")).toBe("Sheet1!$H$10+Sheet1!H11");
    expect(repoint("'Sheet1'!A10")).toBe("'Sheet1'!H10");
    expect(repoint("A10+A12", "s1")).toBe("H10+H12");
  });

  it("leaves a range partly overlapping the source, and anything outside it", () => {
    expect(repoint("SUM(Sheet1!A9:A12)")).toBe("SUM(Sheet1!A9:A12)");
    expect(repoint("Sheet1!B10+A10")).toBe("Sheet1!B10+A10");
  });

  it("turns a reference entirely inside the destination (and not the source) into #REF!", () => {
    expect(repoint("SUM(Sheet1!H10:H11)")).toBe("SUM(Sheet1!#REF!)");
    expect(repoint("H12*2", "s1")).toBe("#REF!*2");
  });

  it("leaves whole rows and columns, 3-D and external references", () => {
    expect(repoint("SUM(Sheet1!A:A)+SUM(Sheet1!10:12)")).toBe("SUM(Sheet1!A:A)+SUM(Sheet1!10:12)");
    expect(repoint("SUM(Sheet1:Sheet2!A10)")).toBe("SUM(Sheet1:Sheet2!A10)");
    expect(repoint("[1]Sheet1!A10")).toBe("[1]Sheet1!A10");
  });

  it("adds, changes or drops the sheet prefix to match the host across sheets", () => {
    const across: RangeMove = { fromSheetId: "s1", from: range(0, 0, 1, 0), toSheetId: "s2", toTop: 4, toLeft: 4 };
    expect(repoint("A1+A2", "s1", across)).toBe("Sheet2!E5+Sheet2!E6"); // added
    expect(repoint("Sheet1!A1", "s3", across)).toBe("Sheet2!E5"); // changed
    expect(repoint("Sheet1!A1", "s2", across)).toBe("E5"); // dropped
  });

  it("does not shift a moved formula otherwise; across sheets its old sheet's references gain the prefix", () => {
    expect(repointFormula("B10*2+A11", "s1", DOWN, nameOf, "s1")).toBe("B10*2+H11");
    const across: RangeMove = { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s3", toTop: 0, toLeft: 0 };
    expect(repointFormula("B1+A1", "s1", across, nameOf, "s3")).toBe("Sheet1!B1+A1");
  });

  it("gives a moved formula's whole rows and columns its old sheet's prefix too, so they keep pointing where they did", () => {
    const across: RangeMove = { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s3", toTop: 0, toLeft: 0 };
    expect(repointFormula("SUM(B:B)+B1", "s1", across, nameOf, "s3")).toBe("SUM(Sheet1!B:B)+Sheet1!B1");
    expect(repointFormula("SUM($2:$3)+SUM(Sheet2!C:C)+SUM(Sheet1:Sheet2!A1)", "s1", across, nameOf, "s3")).toBe("SUM(Sheet1!$2:$3)+SUM(Sheet2!C:C)+SUM(Sheet1:Sheet2!A1)");
    // On its own sheet a moved formula's whole rows and columns stay as written.
    expect(repointFormula("SUM(B:B)", "s1", DOWN, nameOf, "s1")).toBe("SUM(B:B)");
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, 3, { formula: "SUM(B:B)+B1" })])] }), sheet("s2", "Sheet2")]);
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s2", toTop: 4, toLeft: 4 });
    expect(cellOf(m, "s2", 4, 4)?.formula).toBe("SUM(Sheet1!B:B)+Sheet1!B1");
  });
});

describe("moveRange", () => {
  /** Sheet1: A10:A12 = 5, 6, 7 (A11 bold), A12 =A10+A11, B10 =A10*2, H11 "old",
   * A10:A11 is no merge; Sheet2: A1 =SUM(Sheet1!A10:A12), A2 =Sheet1!H11. */
  function book(): SpreadsheetModel {
    return model(
      [
        sheet("s1", "Sheet1", {
          rows: [
            row(9, [value(0, 5), value(1, 10, { formula: "A10*2" })]),
            row(10, [value(0, 6, { styleId: 1 }), value(7, "old")]),
            row(11, [value(0, 11, { formula: "A10+A11" })]),
            row(30, [value(0, "far")]),
          ],
        }),
        sheet("s2", "Sheet2", { rows: [row(0, [value(0, 18, { formula: "SUM(Sheet1!A10:A12)" })]), row(1, [value(0, "old", { formula: "Sheet1!H11" })])] }),
      ],
      { styles: [{}, { font: { bold: true } }] },
    );
  }

  it("overwrites the destination, clears the source to unstyled blanks and re-points every formula", () => {
    const m = book();
    const far = sheetOf(m, "s1").rows[3];
    expect(moveRange(m, DOWN)).toEqual({ ok: true });
    expect(cellOf(m, "s1", 9, 0)).toBeUndefined();
    expect(cellOf(m, "s1", 10, 0)).toBeUndefined();
    expect(cellOf(m, "s1", 9, 7)).toEqual(value(7, 5));
    expect(cellOf(m, "s1", 10, 7)).toEqual(value(7, 6, { styleId: 1 }));
    expect(cellOf(m, "s1", 11, 7)).toEqual(value(7, 11, { formula: "H10+H11" }));
    expect(cellOf(m, "s1", 9, 1)?.formula).toBe("H10*2");
    expect(cellOf(m, "s2", 0, 0)?.formula).toBe("SUM(Sheet1!H10:H12)");
    expect(cellOf(m, "s2", 1, 0)?.formula).toBe("Sheet1!#REF!");
    expect(sheetOf(m, "s1").rows.find((each) => each.index === 30)).toBe(far);
    expect(sheetOf(m, "s1").dirty).toBe(true);
    expect(sheetOf(m, "s2").dirty).toBe(true);
  });

  it("leaves a position styled by its column an unstyled blank, and carries the style a moved cell showed", () => {
    const m = model([sheet("s1", "S", { rows: [row(0, [value(0, 1)])], columns: [{ min: 0, max: 0, styleId: 1 }] })], { styles: [{}, { font: { italic: true } }] });
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s1", toTop: 0, toLeft: 3 });
    expect(cellOf(m, "s1", 0, 0)).toEqual(value(0, null, { styleId: 0 }));
    expect(cellOf(m, "s1", 0, 3)).toEqual(value(3, 1, { styleId: 1 }));
  });

  it("moves a block overlapping its own destination (down one row), reading before writing", () => {
    const m = model([sheet("s1", "S", { rows: [row(0, [value(0, 1)]), row(1, [value(0, 2)]), row(2, [value(0, 3, { formula: "A1+A2" })]), row(5, [value(0, 0, { formula: "SUM(A1:A3)" })])] })]);
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 2, 0), toSheetId: "s1", toTop: 1, toLeft: 0 });
    expect([0, 1, 2, 3].map((r) => cellOf(m, "s1", r, 0)?.value ?? null)).toEqual([null, 1, 2, 3]);
    expect(cellOf(m, "s1", 3, 0)?.formula).toBe("A2+A3");
    expect(cellOf(m, "s1", 5, 0)?.formula).toBe("SUM(A2:A4)");
  });

  it("moves across sheets: cells, merges and arrays to the other sheet, formulas re-pointed with prefixes", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(0, [value(0, "m"), value(2, 3, { formula: "A2*3", arrayRange: range(0, 2, 1, 2) })]), row(1, [value(0, 1), value(2, 3)]), row(4, [value(0, 0, { formula: "A2+C1" })])],
        merges: [range(0, 0, 0, 1)],
      }),
      sheet("s3", "Other Sheet", { rows: [row(9, [value(9, "gone")])], merges: [range(9, 9, 9, 10)] }),
    ]);
    expect(moveRange(m, { fromSheetId: "s1", from: range(0, 0, 1, 2), toSheetId: "s3", toTop: 9, toLeft: 9 })).toEqual({ ok: true });
    expect(sheetOf(m, "s1").merges).toEqual([]);
    expect(sheetOf(m, "s3").merges).toEqual([range(9, 9, 9, 10)]);
    expect(cellOf(m, "s3", 9, 9)?.value).toBe("m");
    expect(cellOf(m, "s3", 9, 11)).toEqual(value(11, 3, { formula: "J11*3", arrayRange: range(9, 11, 10, 11) }));
    expect(cellOf(m, "s1", 4, 0)?.formula).toBe("'Other Sheet'!J11+'Other Sheet'!L10");
  });

  it("refuses a merge or an array crossing either edge, a locked cell, the grid's edge and the cap — changing nothing", () => {
    const cases: [Partial<Sheet>, RangeMove, string][] = [
      [{ merges: [range(0, 0, 0, 1)] }, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s1", toTop: 5, toLeft: 5 }, PASTE_MERGE_MESSAGE],
      [{ merges: [range(5, 5, 6, 5)] }, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s1", toTop: 5, toLeft: 5 }, PASTE_MERGE_MESSAGE],
      [{ rows: [row(0, [value(0, 1, { formula: "1", arrayRange: range(0, 0, 1, 0) })])] }, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s1", toTop: 5, toLeft: 5 }, PARTIAL_ARRAY_MESSAGE],
      [{ protection: { sheet: true } }, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s1", toTop: 5, toLeft: 5 }, PROTECTED_CELL_MESSAGE],
      [{}, { fromSheetId: "s1", from: range(0, 0, 1, 0), toSheetId: "s1", toTop: MAX_ROWS - 1, toLeft: 0 }, PASTE_PAST_EDGE_MESSAGE],
      [{}, { fromSheetId: "s1", from: range(0, 0, 1000, 999), toSheetId: "s1", toTop: 5, toLeft: 5 }, TOO_MANY_TO_CHANGE_MESSAGE],
    ];
    for (const [fields, move, message] of cases) {
      const m = model([sheet("s1", "S", fields)]);
      const sheets = m.sheets;
      expect(moveRefusal(m, move)).toBe(message);
      expect(moveRange(m, move)).toEqual({ ok: false, message });
      expect(m.sheets).toBe(sheets);
    }
    // A merge entirely inside the source moves with it, even over its own destination.
    const own = model([sheet("s1", "S", { merges: [range(0, 0, 1, 0)] })]);
    expect(moveRange(own, { fromSheetId: "s1", from: range(0, 0, 1, 0), toSheetId: "s1", toTop: 1, toLeft: 0 })).toEqual({ ok: true });
    expect(sheetOf(own, "s1").merges).toEqual([range(1, 0, 2, 0)]);
  });

  it("snaps a drag-move's destination inside the grid", () => {
    expect(snappedMove(range(2, 2, 3, 3), { row: 2, col: 2 }, { row: 5, col: 0 })).toEqual({ toTop: 5, toLeft: 0 });
    expect(snappedMove(range(2, 2, 3, 3), { row: 3, col: 3 }, { row: 0, col: 0 })).toEqual({ toTop: 0, toLeft: 0 });
    expect(snappedMove(range(0, 0, 1, 0), { row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 0 })).toEqual({ toTop: MAX_ROWS - 2, toLeft: 0 });
  });
});

describe("moveMapper — retained references", () => {
  const m = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")]);

  it("moves an sqref entry entirely inside the source, drops one entirely inside the destination, keeps partial overlaps", () => {
    const mapper = moveMapper(m, DOWN);
    expect(mapper.range("SHEET1", range(9, 0, 10, 0))).toEqual(range(9, 7, 10, 7));
    expect(mapper.range("Sheet1", range(9, 7, 9, 7))).toBeNull();
    const partial = range(8, 0, 10, 0);
    expect(mapper.range("Sheet1", partial)).toBe(partial);
    expect(mapper.range("Sheet2", range(9, 0, 9, 0))).toEqual(range(9, 0, 9, 0));
  });

  it("leaves indices (drawings, page breaks), and never renames a sheet across sheets", () => {
    expect(moveMapper(m, DOWN).index("Sheet1", "row", 10)).toBe(10);
    const across = moveMapper(m, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s2", toTop: 3, toLeft: 3 });
    const inside = range(0, 0, 0, 0);
    expect(across.range("Sheet1", inside)).toBe(inside);
    expect(across.range("Sheet2", range(3, 3, 3, 3))).toBeNull();
  });
});

// --- SPEC-128's autofilter move adjuster ("Moves") -----------------------------------------

const hiddenRows = (s: Sheet): number[] => s.rows.filter((each) => each.filterHidden).map((each) => each.index);

describe("the autofilter move adjuster (SPEC-128 \"Moves\")", () => {
  const WEST: AutoFilter["columns"] = [{ colId: 0, criteria: { kind: "values", values: ["West"], blanks: false, dateGroups: [] } }];
  /** A1:B5 — Region | Amount over West, East, West, North — filtered to West
   * (rows 3 and 5 filter-hidden), last sorted by Amount descending. */
  const regions = (fields: Partial<Sheet> = {}): Sheet =>
    sheet("s1", "Sheet1", {
      rows: [
        row(0, [value(0, "Region"), value(1, "Amount")]),
        row(1, [value(0, "West"), value(1, 1)]),
        row(2, [value(0, "East"), value(1, 2)], { filterHidden: true }),
        row(3, [value(0, "West"), value(1, 3)]),
        row(4, [value(0, "North"), value(1, 4)], { filterHidden: true }),
      ],
      autoFilter: {
        range: range(0, 0, 4, 1),
        columns: WEST,
        sort: { ref: range(1, 0, 4, 1), keys: [{ col: 1, descending: true }], sourceXml: `<sortState ref="A2:B5"/>` },
        sourceXml: `<autoFilter ref="A1:B5"/>`,
      },
      ...fields,
    });
  it("moves an autofilter entirely inside the source by the offset with its sort record, clears the old rows' filterHidden and reapplies it at its new place", () => {
    const m = model([regions()]);
    expect(moveRange(m, { fromSheetId: "s1", from: range(0, 0, 4, 1), toSheetId: "s1", toTop: 10, toLeft: 3 })).toEqual({ ok: true });
    const s = sheetOf(m, "s1");
    expect(s.autoFilter).toEqual({ range: range(10, 3, 14, 4), columns: WEST, sort: { ref: range(11, 3, 14, 4), keys: [{ col: 4, descending: true }] } });
    expect(hiddenRows(s)).toEqual([12, 14]);
    expect(s.rows.some((each) => each.index < 10)).toBe(false);
    expect(m.structureDirty).toBe(true);
  });

  it("across sheets it replaces the destination's autofilter (whose rows show again) and leaves the source with none", () => {
    const other = sheet("s2", "Sheet2", {
      rows: [row(0, [value(0, "Kind")]), row(1, [value(0, "a")], { filterHidden: true }), row(2, [value(0, "b")])],
      autoFilter: { range: range(0, 0, 2, 0), columns: [{ colId: 0, criteria: { kind: "values", values: ["b"], blanks: false, dateGroups: [] } }], odsName: "__Anonymous_Sheet_DB__1" },
    });
    const m = model([regions(), other], { format: "ods" });
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 4, 1), toSheetId: "s2", toTop: 5, toLeft: 0 });
    expect(sheetOf(m, "s1").autoFilter).toBeUndefined();
    expect(hiddenRows(sheetOf(m, "s1"))).toEqual([]);
    const s2 = sheetOf(m, "s2");
    expect(s2.autoFilter).toEqual({ range: range(5, 0, 9, 1), columns: WEST, sort: { ref: range(6, 0, 9, 1), keys: [{ col: 1, descending: true }] }, odsName: "__Anonymous_Sheet_DB__1" });
    expect(hiddenRows(s2)).toEqual([7, 9]);
  });

  it("leaves an autofilter the source overlaps only partly, and one the destination covers, by identity", () => {
    const m = model([regions()]);
    const { autoFilter } = sheetOf(m, "s1");
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 2, 1), toSheetId: "s1", toTop: 20, toLeft: 0 });
    expect(sheetOf(m, "s1").autoFilter).toBe(autoFilter);
    moveRange(m, { fromSheetId: "s1", from: range(30, 0, 34, 1), toSheetId: "s1", toTop: 0, toLeft: 0 });
    expect(sheetOf(m, "s1").autoFilter).toBe(autoFilter);
    expect(m.structureDirty).toBeUndefined();
  });

  it("moves a worksheet sort record entirely inside the source by the same rule, and leaves a partial one", () => {
    const sortState = { ref: range(1, 5, 9, 6), keys: [{ col: 6, descending: false }], sourceXml: `<sortState ref="F2:G10"/>` };
    const m = model([sheet("s1", "Sheet1", { sortState }), sheet("s2", "Sheet2")]);
    moveRange(m, { fromSheetId: "s1", from: range(0, 5, 9, 6), toSheetId: "s2", toTop: 0, toLeft: 0 });
    expect(sheetOf(m, "s1").sortState).toBeUndefined();
    expect(sheetOf(m, "s2").sortState).toEqual({ ref: range(1, 0, 9, 1), keys: [{ col: 1, descending: false }] });
    expect(m.structureDirty).toBeUndefined();
    const partial = model([sheet("s1", "Sheet1", { sortState })]);
    moveRange(partial, { fromSheetId: "s1", from: range(0, 5, 4, 6), toSheetId: "s1", toTop: 20, toLeft: 5 });
    expect(sheetOf(partial, "s1").sortState).toBe(sortState);
  });

  it("is copy-on-write: the moved objects are new, the old ones untouched", () => {
    const source = regions();
    const snapshot = structuredClone(source);
    const m = model([source]);
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 4, 1), toSheetId: "s1", toTop: 10, toLeft: 3 });
    expect(source).toEqual(snapshot);
  });
});

// --- the registry contract ------------------------------------------------------------------
// Every MoveAdjuster joins this table: a model holding no structure of its
// kind is left untouched by identity, and moving the block holding its
// structure moves it. A later spec appends its row with its adjuster.

/** A model holding one structure of each kind this spec moves, inside A1:B2. */
function withStructure(id: string): SpreadsheetModel {
  const base = sheet("s1", "Sheet1");
  switch (id) {
    case "cells":
      return model([{ ...base, rows: [row(0, [value(0, 1)])] }]);
    case "formulas":
      return model([{ ...base, rows: [row(5, [value(0, 1, { formula: "A1" })])] }]);
    case "merges":
      return model([{ ...base, merges: [range(0, 0, 1, 1)] }]);
    case "retained":
      return model([base]);
    case "autofilter":
      return model([{ ...base, rows: [row(0, [value(0, "Head")]), row(1, [value(0, 1)])], autoFilter: { range: range(0, 0, 1, 1), columns: [] } }]);
    default:
      throw new Error(`no contract fixture for the "${id}" adjuster`);
  }
}

function structureOf(id: string, m: SpreadsheetModel): unknown {
  const s = sheetOf(m, "s1");
  switch (id) {
    case "cells":
      return s.rows.flatMap((each) => each.cells.map((cell) => [each.index, cell.col]));
    case "formulas":
      return getCell(s, 5, 0)?.formula;
    case "merges":
      return s.merges;
    case "autofilter":
      return s.autoFilter?.range;
    default:
      return undefined;
  }
}

describe("MOVE_ADJUSTERS — the registry contract", () => {
  const MOVE: RangeMove = { fromSheetId: "s1", from: range(0, 0, 1, 1), toSheetId: "s1", toTop: 10, toLeft: 3 };

  it("registers this spec's adjusters in order, then SPEC-128's autofilter", () => {
    expect(MOVE_ADJUSTERS.map((adjuster) => adjuster.id)).toEqual(["cells", "formulas", "merges", "retained", "autofilter"]);
  });

  for (const adjuster of MOVE_ADJUSTERS) {
    it(`${adjuster.id}: no structure of its kind → the model untouched by identity`, () => {
      const m = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")]);
      const [sheets, first, second, opaque] = [m.sheets, m.sheets[0], m.sheets[1], m.opaque];
      adjuster.move(m, MOVE);
      adjuster.move(m, { ...MOVE, toSheetId: "s2" });
      expect(m.sheets).toBe(sheets);
      expect(m.sheets[0]).toBe(first);
      expect(m.sheets[1]).toBe(second);
      expect(m.opaque).toBe(opaque);
    });

    if (adjuster.id !== "retained") {
      // retained's own contract is SPEC-124's format halves' specs: a model
      // with no retained source has nothing to move.
      it(`${adjuster.id}: moving the block holding its structure moves it`, () => {
        const m = withStructure(adjuster.id);
        const before = structureOf(adjuster.id, m);
        adjuster.move(m, MOVE);
        const after = structureOf(adjuster.id, m);
        expect(after).not.toEqual(before);
        expect(after).toEqual({ cells: [[10, 3]], formulas: "D11", merges: [range(10, 3, 11, 4)], autofilter: range(10, 3, 11, 4) }[adjuster.id]);
      });
    }
  }
});
