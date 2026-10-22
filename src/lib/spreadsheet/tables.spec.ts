import { describe, expect, it } from "vitest";
import { getCell, type GridRange } from "./address";
import { DEFAULT_THEME } from "../core/theme";
import { clearContents, PROTECTED_CELL_MESSAGE } from "./cells";
import { toOdsFormula } from "./formula/dialect";
import { NAME_PROBLEM_MESSAGES } from "./names";
import {
  convertTableToRange,
  createTable,
  detectHeaderRow,
  expandTableForEntry,
  fillCalculatedColumn,
  headerEditColumn,
  LOCKED_TABLE_MESSAGE,
  qualifyEntryFormula,
  renameTable,
  renameTableColumn,
  RESIZE_TABLE_MESSAGE,
  resizeTable,
  setHeaderRow,
  setTableOptions,
  setTableStyle,
  setTotalsFunction,
  setTotalsRow,
  SUBTOTAL_CODES,
  TABLE_ARRAY_MESSAGE,
  TABLES_OVERLAP_MESSAGE,
  TOTALS_FUNCTIONS,
} from "./tableCommands";
import {
  columnNeedsBrackets,
  escapeColumnName,
  expandStructuredReferences,
  formatStructured,
  parseStructuredText,
  qualifyStructuredReferences,
  removeTableReferences,
  renameColumnInFormula,
  renameTableInFormula,
  resolveStructuredReference,
  structuredFromXlsx,
  structuredToXlsx,
} from "./tables";
import { mergeCells, TABLE_MERGE_MESSAGE } from "./mergeCells";
import type { Cell, Row, Sheet, SheetTable, SpreadsheetModel } from "./model";
import { moveRange } from "./move";
import { PARTIAL_TABLE_MOVE_MESSAGE } from "./nameTableAdjusters";
import { applyStructuralEdit } from "./structure";

// SPEC-133 "Tables follow cells", "Tables and moves": the `tables` adjusters
// (nameTableAdjusters.ts), the partial-table move refusal, the merge refusal
// and clearing a whole table.
describe("tables follow cells", () => {
  const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
  const text = (col: number, value: Cell["value"], formula?: string): Cell => (formula === undefined ? { col, value } : { col, value, formula });
  const row = (index: number, cells: Cell[]): Row => ({ index, cells });

  /** Sales at B2:C5 on Sheet1 (header row 2, data 3–4, totals 5); Sheet2
   * holds formulas naming it. */
  function book(fields: Partial<SheetTable> = {}): SpreadsheetModel {
    const sales: SheetTable = {
      id: "t1",
      name: "Sales",
      ref: range(1, 1, 4, 2),
      headerRow: true,
      totalsRow: true,
      columns: [
        { name: "Price", xlsxId: 1, calculatedFormula: "Sales[@Qty]*$H$1" },
        { name: "Qty", xlsxId: 2, totalsFunction: "custom", totalsFormula: "SUM(Sales[Qty])+$H$1" },
      ],
      style: { name: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
      partName: "xl/tables/table1.xml",
      ...fields,
    };
    const sheet1: Sheet = {
      id: "s1",
      name: "Sheet1",
      visibility: "visible",
      kind: "grid",
      rows: [row(1, [text(1, "Price"), text(2, "Qty")]), row(2, [text(1, 1), text(2, 2)]), row(3, [text(1, 3), text(2, 4)]), row(4, [text(1, "Total")])],
      columns: [],
      defaultColWidthPx: 64,
      defaultRowHeightPt: 15,
      merges: [],
      conditionalFormats: [],
      validations: [],
      tables: [sales],
    };
    const sheet2: Sheet = { ...sheet1, id: "s2", name: "Sheet2", rows: [row(0, [text(0, 0, "SUM(Sales[Qty])"), text(1, 0, "Sales[[#Totals],[Price]]"), text(2, 0, "SUM(Sales)")])], tables: [] };
    return {
      names: [{ name: "Doubled", formula: "SUM(Sales[Price])*2" }],
      dxfs: [],
      kind: "spreadsheet",
      id: "book",
      title: "T",
      format: "xlsx",
      readerVersion: 8,
      sheets: [sheet1, sheet2],
      activeSheetIndex: 0,
      dateSystem: "1900",
      opaque: [],
      styles: [{}],
      defaultStyleId: 0,
    };
  }
  const tableOf = (m: SpreadsheetModel, sheetIndex = 0): SheetTable | undefined => m.sheets[sheetIndex]!.tables?.[0];
  const formulasOn2 = (m: SpreadsheetModel): (string | undefined)[] => m.sheets[1]!.rows[0]!.cells.map((each) => each.formula);
  const edit = (m: SpreadsheetModel, kind: "insertRows" | "deleteRows" | "insertCols" | "deleteCols", at: number, count = 1): void => {
    expect(applyStructuralEdit(m, { kind, sheetId: "s1", at, count })).toEqual({ ok: true });
  };

  it("a row inserted above the table shifts ref; the table and its sheet are dirty", () => {
    const m = book();
    edit(m, "insertRows", 0);
    expect(tableOf(m)).toMatchObject({ ref: range(2, 1, 5, 2), headerRow: true, totalsRow: true, dirty: true, partName: "xl/tables/table1.xml" });
    expect(m.sheets[0]!.dirty).toBe(true);
    edit(m, "insertRows", 1); // at the header row: still a shift
    expect(tableOf(m)!.ref).toEqual(range(3, 1, 6, 2));
  });

  it("the A1 parts of calculated and totals formulas follow the edit like cell formulas", () => {
    const m = book();
    edit(m, "insertRows", 0);
    expect(tableOf(m)!.columns.map((column) => [column.calculatedFormula, column.totalsFormula])).toEqual([
      ["Sales[@Qty]*$H$2", undefined],
      [undefined, "SUM(Sales[Qty])+$H$2"],
    ]);
  });

  it("an insert inside grows ref; new columns are Column{n} with their header cells written", () => {
    const m = book();
    edit(m, "insertRows", 3, 2);
    expect(tableOf(m)!.ref).toEqual(range(1, 1, 6, 2));
    edit(m, "insertCols", 2, 2);
    const grown = tableOf(m)!;
    expect(grown.ref).toEqual(range(1, 1, 6, 4));
    expect(grown.columns.map((column) => [column.name, column.xlsxId])).toEqual([
      ["Price", 1],
      ["Column1", 3],
      ["Column2", 4],
      ["Qty", 2],
    ]);
    expect([2, 3, 4].map((col) => getCell(m.sheets[0]!, 1, col)?.value)).toEqual(["Column1", "Column2", "Qty"]);
  });

  it("deleting the header or totals row turns it off; other overlapping deletions shrink ref", () => {
    const header = book();
    edit(header, "deleteRows", 1);
    expect(tableOf(header)).toMatchObject({ ref: range(1, 1, 3, 2), headerRow: false, totalsRow: true });
    const totals = book();
    edit(totals, "deleteRows", 3, 2);
    expect(tableOf(totals)).toMatchObject({ ref: range(1, 1, 2, 2), headerRow: true, totalsRow: false });
  });

  it("deleting a column removes it and #REF!s references to it; the others stay", () => {
    const m = book();
    edit(m, "deleteCols", 1);
    expect(tableOf(m)).toMatchObject({ ref: range(1, 1, 4, 1), columns: [{ name: "Qty" }] });
    expect(formulasOn2(m)).toEqual(["SUM(Sales[Qty])", "#REF!", "SUM(Sales)"]);
  });

  it("deleting every data row (or every column) deletes the table, its references #REF! everywhere", () => {
    const m = book();
    edit(m, "deleteRows", 2, 2);
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(formulasOn2(m)).toEqual(["SUM(#REF!)", "#REF!", "SUM(#REF!)"]);
    expect(m.names[0]!.formula).toBe("SUM(#REF!)*2");
    const columns = book();
    edit(columns, "deleteCols", 0, 5);
    expect(columns.sheets[0]!.tables).toEqual([]);
  });

  it("moveRange of a block containing the table moves it — onto another sheet too", () => {
    const m = book();
    expect(moveRange(m, { fromSheetId: "s1", from: range(0, 0, 5, 3), toSheetId: "s1", toTop: 10, toLeft: 4 })).toEqual({ ok: true });
    expect(tableOf(m)).toMatchObject({ ref: range(11, 5, 14, 6), dirty: true, partName: "xl/tables/table1.xml" });
    expect(tableOf(m)!.columns[0]!.calculatedFormula).toBe("Sales[@Qty]*$H$1");
    expect(moveRange(m, { fromSheetId: "s1", from: range(11, 5, 14, 6), toSheetId: "s2", toTop: 20, toLeft: 0 })).toEqual({ ok: true });
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(tableOf(m, 1)).toMatchObject({ id: "t1", ref: range(20, 0, 23, 1) });
    expect(formulasOn2(m)).toEqual(["SUM(Sales[Qty])", "Sales[[#Totals],[Price]]", "SUM(Sales)"]);
  });

  it("refuses a move cutting through a table, at its source or its destination", () => {
    const m = book();
    expect(moveRange(m, { fromSheetId: "s1", from: range(1, 1, 2, 2), toSheetId: "s1", toTop: 10, toLeft: 0 })).toEqual({ ok: false, message: PARTIAL_TABLE_MOVE_MESSAGE });
    expect(moveRange(m, { fromSheetId: "s1", from: range(20, 0, 21, 0), toSheetId: "s1", toTop: 3, toLeft: 1 })).toEqual({ ok: false, message: PARTIAL_TABLE_MOVE_MESSAGE });
    expect(tableOf(m)!.ref).toEqual(range(1, 1, 4, 2));
  });

  it("a table overwritten by a move's destination is deleted, its references #REF!", () => {
    const m = book();
    expect(moveRange(m, { fromSheetId: "s1", from: range(20, 0, 25, 3), toSheetId: "s1", toTop: 0, toLeft: 0 })).toEqual({ ok: true });
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(formulasOn2(m)).toEqual(["SUM(#REF!)", "#REF!", "SUM(#REF!)"]);
  });

  it("SPEC-124's merge refuses a range touching a table", () => {
    const m = book();
    expect(mergeCells(m, "s1", [range(0, 0, 1, 1)], "cells")).toEqual({ ok: false, message: TABLE_MERGE_MESSAGE });
    expect(m.sheets[0]!.merges).toEqual([]);
    expect(mergeCells(m, "s1", [range(10, 0, 11, 1)], "cells")).toEqual({ ok: true });
  });

  it("clearing every cell of a table in one command deletes it; clearing part of it does not", () => {
    const m = book();
    clearContents(m, 0, [range(2, 1, 3, 2)]);
    expect(tableOf(m)?.name).toBe("Sales");
    clearContents(m, 0, [range(1, 1, 2, 2), range(3, 0, 4, 5)]);
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(formulasOn2(m)).toEqual(["SUM(#REF!)", "#REF!", "SUM(#REF!)"]);
  });
});
