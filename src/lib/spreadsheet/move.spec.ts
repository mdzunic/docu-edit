// SPEC-125 "Cut and move": every row of the re-point table, #REF! for
// references to overwritten cells, moved formulas not shifted, overlapping
// source and destination, cross-sheet moves, the refusals, moveMapper for
// retained text, and the registry contract every later MoveAdjuster joins.
// SPEC-128's autofilter and SPEC-129's conditionalFormats move adjusters
// have their own blocks.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { getCell } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import { PASTE_MERGE_MESSAGE, PASTE_PAST_EDGE_MESSAGE, TOO_MANY_TO_CHANGE_MESSAGE } from "./clipboard";
import type { CfRule, ConditionalFormat } from "./conditionalFormats";
import type { AutoFilter } from "./filter";
import { MAX_ROWS, type Cell, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { MOVE_ADJUSTERS, moveMapper, moveRange, moveRefusal, repointFormula, snappedMove, type RangeMove } from "./move";
import { PARTIAL_ARRAY_MESSAGE } from "./structure";
import type { DataValidation } from "./validation";

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], validations: [], id, name, visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(sheets: Sheet[], fields: Partial<SpreadsheetModel> = {}): SpreadsheetModel {
  return { names: [], dxfs: [], kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 8, sheets, activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0, ...fields };
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

// --- SPEC-129's conditionalFormats move adjuster ("Moves") ------------------------------------

/** A format on `ranges` with one expression rule. */
const cfFormat = (id: string, ranges: GridRange[], formula: string, priority = 1, fields: Partial<ConditionalFormat> = {}): ConditionalFormat => ({
  id,
  ranges,
  rules: [{ id: `${id}-r`, priority, kind: "expression", formula }],
  ...fields,
});
const formatsOf = (m: SpreadsheetModel, id = "s1"): ConditionalFormat[] => sheetOf(m, id).conditionalFormats;
const formulaOf = (rule: CfRule | undefined): string | undefined => (rule?.kind === "expression" ? rule.formula : undefined);
/** B5:B6 → F5 on Sheet1. */
const B5_B6_TO_F5: RangeMove = { fromSheetId: "s1", from: range(4, 1, 5, 1), toSheetId: "s1", toTop: 4, toLeft: 5 };

describe("the conditionalFormats move adjuster (SPEC-129 \"Moves\")", () => {
  it("cutting B5:B6 out of a format on B5:B20 and pasting at F5 leaves B7:B20 with B7>$Z$1 and creates F5:F6 with F5>$Z$1", () => {
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [cfFormat("f1", [range(4, 1, 19, 1)], "B5>$Z$1", 1, { sourceXml: "<conditionalFormatting/>" })] })]);
    expect(moveRange(m, B5_B6_TO_F5)).toEqual({ ok: true });
    const [source, piece] = formatsOf(m);
    expect(source).toMatchObject({ id: "f1", ranges: [range(6, 1, 19, 1)] });
    expect(source!.sourceXml).toBeUndefined();
    expect(formulaOf(source!.rules[0])).toBe("B7>$Z$1");
    expect(piece!.ranges).toEqual([range(4, 5, 5, 5)]);
    expect(formulaOf(piece!.rules[0])).toBe("F5>$Z$1");
    expect(piece!.id).not.toBe("f1");
    expect(piece!.rules[0]!.id).not.toBe("f1-r");
    // The copy ranks right after the rule it copies.
    expect([source!.rules[0]!.priority, piece!.rules[0]!.priority]).toEqual([1, 2]);
  });

  it("a format wholly inside the source keeps its id and moves, its references following the moved cells", () => {
    const inside = cfFormat("f1", [range(4, 1, 5, 1)], "AND(B5>$Z$1,$B$6<>0)");
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [inside] })]);
    moveRange(m, B5_B6_TO_F5);
    expect(formatsOf(m)).toHaveLength(1);
    expect(formatsOf(m)[0]).toMatchObject({ id: "f1", ranges: [range(4, 5, 5, 5)], rules: [{ id: "f1-r", priority: 1 }] });
    expect(formulaOf(formatsOf(m)[0]!.rules[0])).toBe("AND(F5>$Z$1,$F$6<>0)");
  });

  it("removes the destination rectangle's previous rules: a format inside it goes, one overlapping it keeps the rest, re-anchored", () => {
    const covered = cfFormat("gone", [range(4, 5, 5, 5)], "F5>0", 2);
    const overlapping = cfFormat("over", [range(4, 5, 9, 5)], "F5>$A$1", 3); // F5:F10
    const elsewhere = cfFormat("far", [range(0, 9, 3, 9)], "J1>0", 4);
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [cfFormat("f1", [range(4, 1, 5, 1)], "B5>0"), covered, overlapping, elsewhere] })]);
    moveRange(m, B5_B6_TO_F5);
    const byId = new Map(formatsOf(m).map((format) => [format.id, format]));
    expect(byId.has("gone")).toBe(false);
    expect(byId.get("over")!.ranges).toEqual([range(6, 5, 9, 5)]); // F7:F10
    expect(formulaOf(byId.get("over")!.rules[0])).toBe("F7>$A$1");
    expect(byId.get("far")!.ranges).toBe(elsewhere.ranges);
    // Dense again: f1 (moved whole) 1, over 2, far 3 — gone's rank closed up.
    expect(formatsOf(m).flatMap((format) => format.rules.map((rule) => [format.id, rule.priority]))).toEqual([
      ["f1", 1],
      ["over", 2],
      ["far", 3],
    ]);
  });

  it("moves a block overlapping its own destination (down one row), reading the source before writing", () => {
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [cfFormat("f1", [range(4, 1, 19, 1)], "B5>$Z$1")] })]);
    moveRange(m, { fromSheetId: "s1", from: range(4, 1, 5, 1), toSheetId: "s1", toTop: 5, toLeft: 1 });
    const [source, piece] = formatsOf(m);
    expect(source!.ranges).toEqual([range(7, 1, 19, 1)]); // B8:B20 (B6:B7 overwritten)
    expect(formulaOf(source!.rules[0])).toBe("B8>$Z$1");
    expect(piece!.ranges).toEqual([range(5, 1, 6, 1)]); // B6:B7 — what B5:B6 carried
    expect(formulaOf(piece!.rules[0])).toBe("B6>$Z$1");
  });

  it("across sheets: the arrivals rank lowest, unqualified references keep their sheet, x14 placement and GUIDs follow", () => {
    const bar: CfRule = {
      id: "bar",
      priority: 2,
      kind: "dataBar",
      x14Id: "{00000000-0000-4000-8000-000000000001}",
      bar: { min: { type: "min" }, max: { type: "max" }, color: "#638EC6", gradient: true, axis: "automatic", direction: "context", minLengthPct: 10, maxLengthPct: 90, showValue: true },
    };
    const whole: ConditionalFormat = { ...cfFormat("f1", [range(4, 1, 5, 1)], "B5>$Z$1"), rules: [{ id: "f1-r", priority: 1, kind: "expression", formula: "B5>$Z$1" }, bar] };
    const theirs = cfFormat("g1", [range(0, 0, 0, 0)], "A1>0");
    const m = model([sheet("s1", "Sheet1", { conditionalFormats: [whole, cfFormat("f2", [range(0, 3, 0, 3)], "D1>0", 3)] }), sheet("s2", "Sheet2", { conditionalFormats: [theirs] })]);
    moveRange(m, { ...B5_B6_TO_F5, toSheetId: "s2" });
    expect(formatsOf(m, "s1").map((format) => format.id)).toEqual(["f2"]);
    expect(formatsOf(m, "s1")[0]!.rules[0]!.priority).toBe(1);
    const [kept, arrived] = formatsOf(m, "s2");
    expect(kept!.rules[0]).toBe(theirs.rules[0]);
    expect(arrived).toMatchObject({ id: "f1", ranges: [range(4, 5, 5, 5)] });
    expect(arrived!.rules.map((rule) => rule.priority)).toEqual([2, 3]);
    // B5 moved with the cells; $Z$1 did not, so it keeps naming Sheet1's —
    // a reference to another sheet, which an xlsx writer puts in x14.
    expect(arrived!.rules[0]).toMatchObject({ formula: "F5>Sheet1!$Z$1", x14: true });
    expect(arrived!.rules[1]!.x14Id).toBeUndefined();
  });

  it("re-points another sheet's rule formulas that reference the moved cells, and marks that sheet dirty", () => {
    const other = cfFormat("o1", [range(0, 0, 9, 0)], "A1>Sheet1!$B$5");
    const m = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2", { conditionalFormats: [other] }), sheet("s3", "Other Sheet")]);
    const untouched = m.sheets[2];
    moveRange(m, B5_B6_TO_F5);
    expect(formulaOf(formatsOf(m, "s2")[0]!.rules[0])).toBe("A1>Sheet1!$F$5");
    expect(sheetOf(m, "s2").dirty).toBe(true);
    expect(m.sheets[2]).toBe(untouched);
  });

  it("is copy-on-write: nothing the move started from is mutated, an untouched format keeps its identity and sourceXml", () => {
    // The untouched format ranks first: the copy, ranked after f1's rule, renumbers nothing above it.
    const untouched = cfFormat("f2", [range(30, 0, 31, 0)], "A31>0", 1, { sourceXml: "<b/>" });
    const touched = cfFormat("f1", [range(4, 1, 19, 1)], "B5>$Z$1", 2, { sourceXml: "<a/>" });
    const source = sheet("s1", "Sheet1", { conditionalFormats: [touched, untouched] });
    const snapshot = structuredClone(source);
    const m = model([source]);
    moveRange(m, B5_B6_TO_F5);
    expect(source).toEqual(snapshot);
    expect(formatsOf(m)).toContain(untouched);
  });
});

// --- SPEC-130's validations move adjuster ("Moves") ------------------------------------------

/** A rule on `ranges` — by default a list whose source is $E$2:$E$6. */
const dvRule = (id: string, ranges: GridRange[], fields: Partial<DataValidation> = {}): DataValidation => ({
  id,
  ranges,
  type: "list",
  operator: "between",
  formula1: "$E$2:$E$6",
  allowBlank: true,
  showDropdown: true,
  input: { show: true, title: "Priority", text: "Pick one" },
  error: { show: true, style: "stop", title: "", text: "" },
  ...fields,
});
const validationsOf = (m: SpreadsheetModel, id = "s1"): DataValidation[] => sheetOf(m, id).validations;
/** B2 → D2 on Sheet1. */
const B2_TO_D2: RangeMove = { fromSheetId: "s1", from: range(1, 1, 1, 1), toSheetId: "s1", toTop: 1, toLeft: 3 };

describe("the validations move adjuster (SPEC-130 \"Moves\")", () => {
  it("cutting B2 (a list rule on B2:B10) and pasting at D2 leaves the rule on B3:B10, creates the same rule on D2, and the rule that covered D2 loses it", () => {
    const list = dvRule("v1", [range(1, 1, 9, 1)], { sourceXml: "<dataValidation/>", odsName: "val1" });
    const covering = dvRule("d", [range(0, 3, 4, 3)], { type: "whole", formula1: "1", formula2: "10" }); // D1:D5
    const m = model([sheet("s1", "Sheet1", { validations: [list, covering] })]);
    expect(moveRange(m, B2_TO_D2)).toEqual({ ok: true });
    const [source, trimmed, piece] = validationsOf(m);
    expect(source).toMatchObject({ id: "v1", ranges: [range(2, 1, 9, 1)], formula1: "$E$2:$E$6", odsName: "val1" });
    expect(source!.sourceXml).toBeUndefined();
    expect(trimmed).toMatchObject({ id: "d", ranges: [range(0, 3, 0, 3), range(2, 3, 4, 3)], formula1: "1", formula2: "10" });
    // The same rule, a new identity: a fresh id, no sourceXml or odsName.
    const { id: _id, ranges: _ranges, sourceXml: _source, odsName: _name, ...settings } = list;
    expect(piece).toEqual({ ...settings, id: expect.any(String), ranges: [range(1, 3, 1, 3)] });
    expect(piece!.id).not.toBe("v1");
    expect(validationsOf(m)).toHaveLength(3);
  });

  it("a rule wholly inside the source keeps its id (and its odsName on its own sheet) and moves, its references following the moved cells", () => {
    const inside = dvRule("w", [range(1, 1, 2, 1)], { type: "custom", formula1: "AND(B2>$Z$1,$B$3<>0)", odsName: "val1", sourceXml: "<x/>" });
    const m = model([sheet("s1", "Sheet1", { validations: [inside] })]);
    moveRange(m, { fromSheetId: "s1", from: range(1, 1, 2, 1), toSheetId: "s1", toTop: 1, toLeft: 5 });
    const { sourceXml: _source, ...kept } = inside;
    expect(validationsOf(m)).toEqual([{ ...kept, ranges: [range(1, 5, 2, 5)], formula1: "AND(F2>$Z$1,$F$3<>0)" }]);
    expect("sourceXml" in validationsOf(m)[0]!).toBe(false);
  });

  it("the list source $E$2:$E$6 follows only if those cells were moved — on every sheet, marking a sheet whose rule changed dirty", () => {
    const own = dvRule("v1", [range(1, 1, 9, 1)]);
    const other = dvRule("o", [range(0, 0, 0, 0)], { formula1: "Sheet1!$E$2:$E$6", x14: true });
    const m = model([sheet("s1", "Sheet1", { validations: [own] }), sheet("s2", "Sheet2", { validations: [other] }), sheet("s3", "Other Sheet")]);
    const untouched = m.sheets[2];
    // Only E2:E3 moves: the source is partly moved, so nothing follows.
    moveRange(m, { fromSheetId: "s1", from: range(1, 4, 2, 4), toSheetId: "s1", toTop: 1, toLeft: 7 });
    expect(validationsOf(m)[0]).toBe(own);
    expect(validationsOf(m, "s2")[0]).toBe(other);
    // All of E2:E6 moves to G2: both follow.
    moveRange(m, { fromSheetId: "s1", from: range(1, 4, 5, 4), toSheetId: "s1", toTop: 1, toLeft: 6 });
    expect(validationsOf(m)[0]).toMatchObject({ id: "v1", ranges: [range(1, 1, 9, 1)], formula1: "$G$2:$G$6" });
    expect(validationsOf(m, "s2")[0]).toMatchObject({ id: "o", formula1: "Sheet1!$G$2:$G$6", x14: true });
    expect(sheetOf(m, "s2").dirty).toBe(true);
    expect(m.sheets[2]).toBe(untouched);
  });

  it("across sheets: a whole rule arrives with its id and without its odsName, a piece rebased; unqualified references keep their sheet, placed in x14", () => {
    const whole = dvRule("w", [range(1, 1, 2, 1)], { odsName: "val1" }); // B2:B3
    const partly = dvRule("p", [range(1, 2, 9, 2)], { type: "custom", formula1: "C2>$Z$1" }); // C2:C10
    const theirs = dvRule("t", [range(4, 3, 19, 3)], { type: "custom", formula1: 'D5<>""' }); // Sheet2!D5:D20
    const m = model([sheet("s1", "Sheet1", { validations: [whole, partly] }), sheet("s2", "Sheet2", { validations: [theirs] })]);
    // B2:C3 → Sheet2!D5:E6.
    moveRange(m, { fromSheetId: "s1", from: range(1, 1, 2, 2), toSheetId: "s2", toTop: 4, toLeft: 3 });
    expect(validationsOf(m, "s1")).toEqual([{ ...partly, ranges: [range(3, 2, 9, 2)], formula1: "C4>$Z$1" }]);
    const [kept, arrived, piece] = validationsOf(m, "s2");
    expect(kept).toEqual({ ...theirs, ranges: [range(6, 3, 19, 3)], formula1: 'D7<>""' });
    const { odsName: _name, ...withoutName } = whole;
    expect(arrived).toEqual({ ...withoutName, ranges: [range(4, 3, 5, 3)], formula1: "Sheet1!$E$2:$E$6", x14: true });
    expect(piece).toMatchObject({ type: "custom", ranges: [range(4, 4, 5, 4)], formula1: "E5>Sheet1!$Z$1", x14: true });
    expect(piece!.id).not.toBe("p");
  });

  it("moves a block overlapping its own destination (down one row), reading the source before writing", () => {
    const m = model([sheet("s1", "Sheet1", { validations: [dvRule("v1", [range(4, 1, 19, 1)], { type: "custom", formula1: "B5>$Z$1" })] })]);
    moveRange(m, { fromSheetId: "s1", from: range(4, 1, 5, 1), toSheetId: "s1", toTop: 5, toLeft: 1 });
    const [source, piece] = validationsOf(m);
    expect(source).toMatchObject({ id: "v1", ranges: [range(7, 1, 19, 1)], formula1: "B8>$Z$1" }); // B8:B20 (B6:B7 overwritten)
    expect(piece).toMatchObject({ ranges: [range(5, 1, 6, 1)], formula1: "B6>$Z$1" }); // B6:B7 — what B5:B6 carried
  });

  it("never merges a moved rule into an identical one at the destination", () => {
    const identical = dvRule("same", [range(2, 3, 9, 3)]); // D3:D10, the same settings as v1
    const m = model([sheet("s1", "Sheet1", { validations: [dvRule("v1", [range(1, 1, 9, 1)]), identical] })]);
    moveRange(m, B2_TO_D2);
    const rules = validationsOf(m);
    expect(rules).toHaveLength(3);
    expect(rules[1]).toBe(identical);
    expect(rules[2]!.ranges).toEqual([range(1, 3, 1, 3)]);
  });

  it("is copy-on-write: nothing the move started from is mutated; a rule the move misses keeps its identity and sourceXml", () => {
    const untouched = dvRule("far", [range(30, 0, 31, 0)], { sourceXml: "<b/>" });
    const touched = dvRule("v1", [range(1, 1, 9, 1)], { sourceXml: "<a/>" });
    const source = sheet("s1", "Sheet1", { validations: [touched, untouched] });
    const snapshot = structuredClone(source);
    const m = model([source]);
    moveRange(m, B2_TO_D2);
    expect(source).toEqual(snapshot);
    expect(validationsOf(m)).toContain(untouched);
    // A move touching no rule leaves the array itself alone.
    const quiet = model([sheet("s1", "Sheet1", { validations: [untouched] })]);
    const rules = validationsOf(quiet);
    moveRange(quiet, B2_TO_D2);
    expect(validationsOf(quiet)).toBe(rules);
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
    case "conditionalFormats":
      return model([{ ...base, conditionalFormats: [cfFormat("f1", [range(0, 0, 1, 1)], "A1>0")] }]);
    case "validations":
      return model([{ ...base, validations: [dvRule("v1", [range(0, 0, 1, 1)])] }]);
    case "charts":
      return model([
        {
          ...base,
          drawings: [
            {
              id: "dr1",
              name: "Chart 1",
              anchor: { kind: "twoCell", editAs: "twoCell", from: { col: 4, colOffEmu: 0, row: 0, rowOffEmu: 0 }, to: { col: 8, colOffEmu: 0, row: 3, rowOffEmu: 0 } },
              content: {
                kind: "chart",
                visual: { chartType: "line", categories: ["1", "2"], series: [{ name: "S", values: [1, 2], color: "#000000" }], options: { legend: "none", dataLabels: false, gridlines: false } },
                refs: { series: [{ values: "Sheet1!$A$1:$A$2" }], plotVisibleOnly: true },
              },
            },
          ],
        },
      ]);
    case "names":
      return model([base], { names: [{ name: "Block", formula: "Sheet1!$A$1:$B$2" }] });
    case "tables":
      return model([
        {
          ...base,
          tables: [
            {
              id: "t1",
              name: "Sales",
              ref: range(0, 0, 1, 1),
              headerRow: true,
              totalsRow: false,
              columns: [{ name: "A" }, { name: "B" }],
              style: { showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
            },
          ],
        },
      ]);
    case "notes":
      return model([{ ...base, notes: [{ noteId: "n-A1", ref: { row: 0, col: 0 }, author: "Ada", text: "Check" }] }]);
    case "hyperlinks":
      return model([{ ...base, hyperlinks: [{ linkId: "h-A1", range: range(0, 0, 1, 1), target: { kind: "url", href: "https://example.com" } }] }]);
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
    case "conditionalFormats":
      return s.conditionalFormats[0]?.ranges;
    case "validations":
      return s.validations[0]?.ranges;
    case "charts": {
      const content = s.drawings?.[0]?.content;
      return content?.kind === "chart" ? content.refs?.series[0]?.values : undefined;
    }
    case "names":
      return m.names[0]?.formula;
    case "notes":
      return s.notes?.[0]?.ref;
    case "hyperlinks":
      return s.hyperlinks?.[0]?.range;
    case "tables":
      return s.tables?.[0]?.ref;
    default:
      return undefined;
  }
}

describe("MOVE_ADJUSTERS — the registry contract", () => {
  const MOVE: RangeMove = { fromSheetId: "s1", from: range(0, 0, 1, 1), toSheetId: "s1", toTop: 10, toLeft: 3 };

  it("registers this spec's adjusters in order, then SPEC-128's autofilter, SPEC-129's conditionalFormats, SPEC-130's validations, SPEC-132's charts, SPEC-133's names and tables and SPEC-135's notes and hyperlinks", () => {
    expect(MOVE_ADJUSTERS.map((adjuster) => adjuster.id)).toEqual(["cells", "formulas", "merges", "retained", "autofilter", "conditionalFormats", "validations", "charts", "names", "tables", "notes", "hyperlinks"]);
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
        expect(after).toEqual(
          {
            cells: [[10, 3]],
            formulas: "D11",
            merges: [range(10, 3, 11, 4)],
            autofilter: range(10, 3, 11, 4),
            conditionalFormats: [range(10, 3, 11, 4)],
            validations: [range(10, 3, 11, 4)],
            charts: "Sheet1!$D$11:$D$12",
            names: "Sheet1!$D$11:$E$12",
            tables: range(10, 3, 11, 4),
            notes: { row: 10, col: 3 },
            hyperlinks: range(10, 3, 11, 4),
          }[adjuster.id],
        );
      });
    }
  }
});
