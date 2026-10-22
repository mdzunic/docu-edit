import { describe, expect, it } from "vitest";
import { getCell, type GridRange } from "./address";
import { clearContents } from "./cells";
import { TABLE_ARRAY_MESSAGE, TABLES_OVERLAP_MESSAGE, TOTALS_FUNCTIONS } from "./tableCommands";
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

const gridRange = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const cellOf = (col: number, value: Cell["value"], formula?: string): Cell => (formula === undefined ? { col, value } : { col, value, formula });
const rowOf = (index: number, cells: Cell[]): Row => ({ index, cells });
const tableOf = (m: SpreadsheetModel, sheetIndex = 0): SheetTable | undefined => m.sheets[sheetIndex]!.tables?.[0];
const formulasOn2 = (m: SpreadsheetModel): (string | undefined)[] => m.sheets[1]!.rows[0]!.cells.map((each) => each.formula);
const editSheet1 = (m: SpreadsheetModel, kind: "insertRows" | "deleteRows" | "insertCols" | "deleteCols", at: number, count = 1): void => {
  expect(applyStructuralEdit(m, { kind, sheetId: "s1", at, count })).toEqual({ ok: true });
};
const hostAt = (row: number, col = 5) => ({ sheetId: "s1", row, col });

// SPEC-133 "Tables follow cells", "Tables and moves": the `tables` adjusters
// (nameTableAdjusters.ts), the partial-table move refusal, the merge refusal
// and clearing a whole table.
describe("tables follow cells", () => {
  /** Sales at B2:C5 on Sheet1 (header row 2, data 3–4, totals 5); Sheet2
   * holds formulas naming it. */
  function book(fields: Partial<SheetTable> = {}): SpreadsheetModel {
    const sales: SheetTable = {
      id: "t1",
      name: "Sales",
      ref: gridRange(1, 1, 4, 2),
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
      rows: [rowOf(1, [cellOf(1, "Price"), cellOf(2, "Qty")]), rowOf(2, [cellOf(1, 1), cellOf(2, 2)]), rowOf(3, [cellOf(1, 3), cellOf(2, 4)]), rowOf(4, [cellOf(1, "Total")])],
      columns: [],
      defaultColWidthPx: 64,
      defaultRowHeightPt: 15,
      merges: [],
      conditionalFormats: [],
      validations: [],
      tables: [sales],
    };
    const sheet2: Sheet = { ...sheet1, id: "s2", name: "Sheet2", rows: [rowOf(0, [cellOf(0, 0, "SUM(Sales[Qty])"), cellOf(1, 0, "Sales[[#Totals],[Price]]"), cellOf(2, 0, "SUM(Sales)")])], tables: [] };
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

  it("a row inserted above the table shifts ref; the table and its sheet are dirty", () => {
    const m = book();
    editSheet1(m, "insertRows", 0);
    expect(tableOf(m)).toMatchObject({ ref: gridRange(2, 1, 5, 2), headerRow: true, totalsRow: true, dirty: true, partName: "xl/tables/table1.xml" });
    expect(m.sheets[0]!.dirty).toBe(true);
    editSheet1(m, "insertRows", 1); // at the header row: still a shift
    expect(tableOf(m)!.ref).toEqual(gridRange(3, 1, 6, 2));
  });

  it("the A1 parts of calculated and totals formulas follow the edit like cell formulas", () => {
    const m = book();
    editSheet1(m, "insertRows", 0);
    expect(tableOf(m)!.columns.map((column) => [column.calculatedFormula, column.totalsFormula])).toEqual([
      ["Sales[@Qty]*$H$2", undefined],
      [undefined, "SUM(Sales[Qty])+$H$2"],
    ]);
  });

  it("an insert inside grows ref; new columns are Column{n} with their header cells written", () => {
    const m = book();
    editSheet1(m, "insertRows", 3, 2);
    expect(tableOf(m)!.ref).toEqual(gridRange(1, 1, 6, 2));
    editSheet1(m, "insertCols", 2, 2);
    const grown = tableOf(m)!;
    expect(grown.ref).toEqual(gridRange(1, 1, 6, 4));
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
    editSheet1(header, "deleteRows", 1);
    expect(tableOf(header)).toMatchObject({ ref: gridRange(1, 1, 3, 2), headerRow: false, totalsRow: true });
    const totals = book();
    editSheet1(totals, "deleteRows", 3, 2);
    expect(tableOf(totals)).toMatchObject({ ref: gridRange(1, 1, 2, 2), headerRow: true, totalsRow: false });
  });

  it("deleting a column removes it and #REF!s references to it; the others stay", () => {
    const m = book();
    editSheet1(m, "deleteCols", 1);
    expect(tableOf(m)).toMatchObject({ ref: gridRange(1, 1, 4, 1), columns: [{ name: "Qty" }] });
    expect(formulasOn2(m)).toEqual(["SUM(Sales[Qty])", "#REF!", "SUM(Sales)"]);
  });

  it("deleting every data row (or every column) deletes the table, its references #REF! everywhere", () => {
    const m = book();
    editSheet1(m, "deleteRows", 2, 2);
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(formulasOn2(m)).toEqual(["SUM(#REF!)", "#REF!", "SUM(#REF!)"]);
    expect(m.names[0]!.formula).toBe("SUM(#REF!)*2");
    const columns = book();
    editSheet1(columns, "deleteCols", 0, 5);
    expect(columns.sheets[0]!.tables).toEqual([]);
  });

  it("moveRange of a block containing the table moves it — onto another sheet too", () => {
    const m = book();
    expect(moveRange(m, { fromSheetId: "s1", from: gridRange(0, 0, 5, 3), toSheetId: "s1", toTop: 10, toLeft: 4 })).toEqual({ ok: true });
    expect(tableOf(m)).toMatchObject({ ref: gridRange(11, 5, 14, 6), dirty: true, partName: "xl/tables/table1.xml" });
    expect(tableOf(m)!.columns[0]!.calculatedFormula).toBe("Sales[@Qty]*$H$1");
    expect(moveRange(m, { fromSheetId: "s1", from: gridRange(11, 5, 14, 6), toSheetId: "s2", toTop: 20, toLeft: 0 })).toEqual({ ok: true });
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(tableOf(m, 1)).toMatchObject({ id: "t1", ref: gridRange(20, 0, 23, 1) });
    expect(formulasOn2(m)).toEqual(["SUM(Sales[Qty])", "Sales[[#Totals],[Price]]", "SUM(Sales)"]);
  });

  it("refuses a move cutting through a table, at its source or its destination", () => {
    const m = book();
    expect(moveRange(m, { fromSheetId: "s1", from: gridRange(1, 1, 2, 2), toSheetId: "s1", toTop: 10, toLeft: 0 })).toEqual({ ok: false, message: PARTIAL_TABLE_MOVE_MESSAGE });
    expect(moveRange(m, { fromSheetId: "s1", from: gridRange(20, 0, 21, 0), toSheetId: "s1", toTop: 3, toLeft: 1 })).toEqual({ ok: false, message: PARTIAL_TABLE_MOVE_MESSAGE });
    expect(tableOf(m)!.ref).toEqual(gridRange(1, 1, 4, 2));
  });

  it("a table overwritten by a move's destination is deleted, its references #REF!", () => {
    const m = book();
    expect(moveRange(m, { fromSheetId: "s1", from: gridRange(20, 0, 25, 3), toSheetId: "s1", toTop: 0, toLeft: 0 })).toEqual({ ok: true });
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(formulasOn2(m)).toEqual(["SUM(#REF!)", "#REF!", "SUM(#REF!)"]);
  });

  it("SPEC-124's merge refuses a range touching a table", () => {
    const m = book();
    expect(mergeCells(m, "s1", [gridRange(0, 0, 1, 1)], "cells")).toEqual({ ok: false, message: TABLE_MERGE_MESSAGE });
    expect(m.sheets[0]!.merges).toEqual([]);
    expect(mergeCells(m, "s1", [gridRange(10, 0, 11, 1)], "cells")).toEqual({ ok: true });
  });

  it("clearing every cell of a table in one command deletes it; clearing part of it does not", () => {
    const m = book();
    clearContents(m, 0, [gridRange(2, 1, 3, 2)]);
    expect(tableOf(m)?.name).toBe("Sales");
    clearContents(m, 0, [gridRange(1, 1, 2, 2), gridRange(3, 0, 4, 5)]);
    expect(m.sheets[0]!.tables).toEqual([]);
    expect(formulasOn2(m)).toEqual(["SUM(#REF!)", "#REF!", "SUM(#REF!)"]);
  });
});

describe("structured references resolve (\"Structured references\")", () => {
  // Sales at A1:D10 on s1: header row 1, data 2–9, totals row 10.
  const sales: SheetTable = {
    id: "t",
    name: "Sales",
    ref: { top: 0, left: 0, bottom: 9, right: 3 },
    headerRow: true,
    totalsRow: true,
    columns: [{ name: "Region" }, { name: "Price" }, { name: "Amount" }, { name: "a.b" }],
    style: { name: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
  };
  const book = { sheets: [{ id: "s1", name: "Data", tables: [sales] } as unknown as Sheet] };
  const area = (text: string, host = hostAt(4)) => {
    const open = text.indexOf("[");
    const value = resolveStructuredReference(book, open > 0 ? { table: text.slice(0, open), text } : { text }, host);
    return "error" in value ? value.error : (({ top, left, bottom, right }) => [top, left, bottom, right])(value.areas[0]!);
  };

  it.each([
    ["Sales", [1, 0, 8, 3]],
    ["Sales[#Data]", [1, 0, 8, 3]],
    ["Sales[#All]", [0, 0, 9, 3]],
    ["Sales[#Headers]", [0, 0, 0, 3]],
    ["Sales[#Totals]", [9, 0, 9, 3]],
    ["Sales[Amount]", [1, 2, 8, 2]],
    ["Sales[[Amount]]", [1, 2, 8, 2]],
    ["Sales[[#Headers],[Amount]]", [0, 2, 0, 2]],
    ["Sales[[#Totals],[Amount]]", [9, 2, 9, 2]],
    ["Sales[[#All],[Amount]]", [0, 2, 9, 2]],
    ["Sales[[#Headers],[#Data],[Amount]]", [0, 2, 8, 2]],
    ["Sales[[#Data],[#Totals],[Amount]]", [1, 2, 9, 2]],
    ["Sales[[Price]:[a.b]]", [1, 1, 8, 3]],
    ["Sales[[a.b]:[Price]]", [1, 1, 8, 3]],
    ["Sales[[#Headers],[#Totals]]", "#REF!"],
    ["Sales[Missing]", "#REF!"],
    ["Nope[Amount]", "#REF!"],
  ] as const)("%s → %j", (text, expected) => {
    expect(area(text)).toEqual(expected);
  });

  it("this-row forms take the host's row inside the data body, #VALUE! outside it", () => {
    expect(area("Sales[@Amount]", hostAt(4, 2))).toEqual([4, 2, 4, 2]);
    expect(area("Sales[@]", hostAt(4, 2))).toEqual([4, 0, 4, 3]);
    expect(area("Sales[@[Price]:[Amount]]", hostAt(4, 2))).toEqual([4, 1, 4, 2]);
    expect(area("[@Amount]", hostAt(4, 2))).toEqual([4, 2, 4, 2]);
    expect(area("Sales[@Amount]", hostAt(0, 2))).toBe("#VALUE!");
    expect(area("Sales[@Amount]", hostAt(9, 2))).toBe("#VALUE!");
    expect(area("Sales[@Amount]", hostAt(4, 5))).toBe("#VALUE!");
  });

  it("a missing header or totals row is #REF!", () => {
    const bare = { sheets: [{ id: "s1", name: "Data", tables: [{ ...sales, headerRow: false, totalsRow: false }] } as unknown as Sheet] };
    expect(resolveStructuredReference(bare, { table: "Sales", text: "Sales[#Headers]" }, hostAt(4))).toEqual({ error: "#REF!" });
    expect(resolveStructuredReference(bare, { table: "Sales", text: "Sales[#Totals]" }, hostAt(4))).toEqual({ error: "#REF!" });
  });

  it("maps @ to and from the file's [#This Row] form, leaving every other token as written", () => {
    const pairs: [string, string][] = [
      ["Sales[@Amount]*2", "Sales[[#This Row],[Amount]]*2"],
      ["Sales[@]", "Sales[#This Row]"],
      ["SUM(Sales[@[Price]:[Amount]])", "SUM(Sales[[#This Row],[Price]:[Amount]])"],
      ["Sales[@['#Units]]", "Sales[[#This Row],['#Units]]"],
      ["SUM(Sales[Amount]) + A1", "SUM(Sales[Amount]) + A1"],
    ];
    for (const [canonical, wire] of pairs) {
      expect(structuredToXlsx(canonical)).toBe(wire);
      expect(structuredFromXlsx(wire)).toBe(canonical);
    }
  });
});

describe("token-level table rewrites", () => {
  it("renames a table, a column, and turns a removed table or column into #REF!, touching only its tokens", () => {
    expect(renameTableInFormula("SUM(Sales[Amount]) + sales[@Price] + Sales + Other[X]", "Sales", "Orders")).toBe("SUM(Orders[Amount]) + Orders[@Price] + Orders + Other[X]");
    expect(renameColumnInFormula("Sales[Amount] + Sales[[Price]:[amount]] + Other[Amount]", "Sales", "Amount", "a.b")).toBe("Sales[[a.b]] + Sales[[Price]:[a.b]] + Other[Amount]");
    expect(removeTableReferences("SUM(Sales[Amount]) + Sales + Other[X]", "Sales")).toBe("SUM(#REF!) + #REF! + Other[X]");
    expect(removeTableReferences("Sales[Amount] + Sales[Price]", "Sales", "amount")).toBe("#REF! + Sales[Price]");
    expect(renameTableInFormula("A1 +  B2", "Sales", "Orders")).toBe("A1 +  B2");
  });
});

describe("the grammar, escapes, qualification and ODS expansion", () => {
  it("parses and prints the canonical spelling, escaping [ ] # ' and bracketing special column names", () => {
    expect(parseStructuredText("Sales[[#This Row],['#Units]]")).toEqual({ table: "Sales", items: [], thisRow: true, columns: { first: "#Units" } });
    expect(formatStructured({ table: "Sales", items: [], thisRow: true, columns: { first: "#Units" } })).toBe("Sales[@['#Units]]");
    expect(formatStructured({ table: "T", items: [], thisRow: false, columns: { first: "a.b" } })).toBe("T[[a.b]]");
    expect(formatStructured({ table: "T", items: [], thisRow: false, columns: { first: "Unit Price" } })).toBe("T[Unit Price]");
    expect(columnNeedsBrackets("Unit Price")).toBe(false);
    expect(["a,b", "a:b", "a.b", "a[b", "a]b", "a#b", "a'b", 'a"b', "a{b", "a}b", "a$b", "a^b", "a&b", "a*b", "a+b", "a=b", "a-b", "a>b", "a<b", "a/b", "a\tb", "a\nb"].every(columnNeedsBrackets)).toBe(true);
    expect(escapeColumnName("[x]#'")).toBe("'[x']'#''");
    expect(parseStructuredText("Sales[[#Headers],[#Data],[Amount]]")).toEqual({ table: "Sales", items: ["#Headers", "#Data"], thisRow: false, columns: { first: "Amount" } });
  });

  it("qualifies a formula's unnamed references with the host table, @ for this-row", () => {
    expect(qualifyStructuredReferences("[@Price]*2+[Qty]+Other[X]", "Sales")).toBe("Sales[@Price]*2+Sales[Qty]+Other[X]");
    expect(qualifyStructuredReferences("[[#This Row],[Price]]", "Sales")).toBe("Sales[@Price]");
  });

  it("expands every structured reference to its A1 equivalent for ODS (\"Canonical and wire forms\")", () => {
    const sales: SheetTable = {
      id: "t",
      name: "Sales",
      ref: { top: 0, left: 0, bottom: 9, right: 3 },
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Region" }, { name: "Price" }, { name: "Amount" }, { name: "Qty" }],
      style: { showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    };
    const book = { names: [], sheets: [{ id: "s1", name: "Data", tables: [sales] } as unknown as Sheet] };
    const host = { sheetId: "s1", row: 6, col: 2 };
    expect(expandStructuredReferences(book, "Sales[@Amount]", host)).toBe("$C7");
    expect(expandStructuredReferences(book, "Sales[@]", host)).toBe("$A7:$D7");
    expect(expandStructuredReferences(book, "Sales[Amount]", host)).toBe("$C$2:$C$10");
    expect(expandStructuredReferences(book, "SUBTOTAL(109,Sales[Amount])", host)).toBe("SUBTOTAL(109,$C$2:$C$10)");
  });
});

describe("the table commands' pinned texts", () => {
  it("refuses with the spec's messages and offers the nine totals functions", () => {
    expect(TABLES_OVERLAP_MESSAGE).toBe("Tables can't overlap");
    expect(TABLE_ARRAY_MESSAGE).toBe("Tables can't contain multi-cell array formulas");
    expect(TOTALS_FUNCTIONS).toHaveLength(9);
    expect(TOTALS_FUNCTIONS).toEqual(expect.arrayContaining(["none", "average", "count", "countNums", "max", "min", "sum", "stdDev", "var"]));
  });
});
