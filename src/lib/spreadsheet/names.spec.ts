// SPEC-133 "Name validity", "Resolution": the validity table row by row,
// uniqueTableName, resolution order, the relative wrap, the nesting cap, the
// `names` adjusters (structural edit and move) and SPEC-121's sheet commands
// reaching names and tables.
import { describe, expect, it } from "vitest";
import { getCell, type GridRange } from "./address";
import { recalculate } from "./formula/recalc";
import type { Cell, DefinedName, Row, Sheet, SheetTable, SpreadsheetModel } from "./model";
import { moveRange } from "./move";
import { NAME_PROBLEM_MESSAGES, rebaseRelative, resolveDefinedName, uniqueTableName, validateName, validateTableName } from "./names";
import { deleteSheet, duplicateSheet, renameSheet } from "./sheets";
import { applyStructuralEdit } from "./structure";

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], validations: [], id, name, visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(sheets: Sheet[], names: DefinedName[] = []): SpreadsheetModel {
  return { names, dxfs: [], kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 8, sheets, activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0 };
}

const cell = (col: number, value: Cell["value"], formula?: string): Cell => (formula === undefined ? { col, value } : { col, value, formula });
const row = (index: number, cells: Cell[]): Row => ({ index, cells });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

function table(name: string, ref: GridRange, fields: Partial<SheetTable> = {}): SheetTable {
  return {
    id: `id-${name}`,
    name,
    ref,
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "Price" }, { name: "Qty" }],
    style: { showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    ...fields,
  };
}

const valueAt = (m: SpreadsheetModel, sheetIndex: number, r: number, c: number): Cell["value"] | undefined => getCell(m.sheets[sheetIndex]!, r, c)?.value;

describe("validateName — the validity table", () => {
  const book = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2", { tables: [table("Sales", range(0, 0, 3, 1))] })], [
    { name: "TaxRate", formula: "0.2" },
    { name: "Local", scopeSheetId: "s1", formula: "1" },
  ]);

  it.each([
    ["", "empty"],
    ["x".repeat(256), "tooLong"],
    ["B12", "looksLikeReference"],
    ["$xfd$1", "looksLikeReference"],
    ["R", "looksLikeReference"],
    ["c", "looksLikeReference"],
    ["R2", "looksLikeReference"],
    ["C3", "looksLikeReference"],
    ["r1c1", "looksLikeReference"],
    ["R1C1", "looksLikeReference"],
    ["1abc", "invalidCharacters"],
    ["has space", "invalidCharacters"],
    ["a-b", "invalidCharacters"],
    ["_xlnm.Print_Area", "reserved"],
    ["taxrate", "duplicate"],
    ["sales", "tableConflict"],
  ])("%s → %s", (name, problem) => {
    expect(validateName(book, name)).toBe(problem);
    expect(NAME_PROBLEM_MESSAGES[problem as keyof typeof NAME_PROBLEM_MESSAGES]).toBeTruthy();
  });

  it("accepts names outside the grid and ordinary ones; uniqueness is per scope", () => {
    expect(validateName(book, "XFE1")).toBeNull();
    expect(validateName(book, "Sales2024")).toBeNull();
    expect(validateName(book, "x".repeat(255))).toBeNull();
    expect(validateName(book, "LOCAL", "s1")).toBe("duplicate");
    expect(validateName(book, "LOCAL")).toBeNull();
    expect(validateName(book, "TaxRate", "s1")).toBeNull();
    expect(validateName(book, "TaxRate", undefined, book.names[0])).toBeNull();
  });

  it("tables are unique against names of every scope and other tables", () => {
    expect(validateTableName(book, "local")).toBe("tableConflict");
    expect(validateTableName(book, "SALES")).toBe("tableConflict");
    expect(validateTableName(book, "Sales", "id-Sales")).toBeNull();
  });

  it("uniqueTableName: Table{n} for a new table, <name>_<n> (n ≥ 2) for a copy", () => {
    const tables = model([sheet("s1", "Sheet1", { tables: [table("Table1", range(0, 0, 2, 1)), table("Sales_2", range(5, 0, 7, 1))] })], [{ name: "Table2", formula: "1" }]);
    expect(uniqueTableName(tables)).toBe("Table3");
    expect(uniqueTableName(tables, "Sales")).toBe("Sales_3");
    expect(uniqueTableName(tables, "Sales", ["Sales_3"])).toBe("Sales_4");
  });
});

describe("resolveDefinedName — resolution order", () => {
  const book = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")], [
    { name: "X", formula: "1" },
    { name: "X", scopeSheetId: "s2", formula: "2" },
    { name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Sheet1!$A$1:$B$2" },
    { name: "_xlfn.FUTURE", formula: "#NAME?", hidden: true },
    { name: "Orphan", formula: "3", sourceAttrs: 'localSheetId="9"' },
  ]);

  it("looks in the host sheet's scope, then the workbook's; a qualified name in that sheet's", () => {
    expect(resolveDefinedName(book, "x", undefined, "s1")?.formula).toBe("1");
    expect(resolveDefinedName(book, "X", undefined, "s2")?.formula).toBe("2");
    expect(resolveDefinedName(book, "X", "Sheet2", "s1")?.formula).toBe("2");
    expect(resolveDefinedName(book, "X", "Sheet1", "s2")?.formula).toBe("1");
    expect(resolveDefinedName(book, "X", "Nowhere", "s1")).toBeUndefined();
  });

  it("answers a reserved name's short form within its sheet; placeholders and orphans never resolve", () => {
    expect(resolveDefinedName(book, "Print_Area", undefined, "s1")?.name).toBe("_xlnm.Print_Area");
    expect(resolveDefinedName(book, "Print_Area", undefined, "s2")).toBeUndefined();
    expect(resolveDefinedName(book, "_xlfn.FUTURE", undefined, "s1")).toBeUndefined();
    expect(resolveDefinedName(book, "Orphan", undefined, "s1")).toBeUndefined();
  });
});

describe("relative names — the wrap", () => {
  it("rebaseRelative moves relative parts modulo the grid, absolute ones not", () => {
    expect(rebaseRelative("Sheet1!B5", -4, -1)).toBe("Sheet1!A1");
    expect(rebaseRelative("Sheet1!A1", -1, 0)).toBe("Sheet1!A1048576");
    expect(rebaseRelative("Sheet1!$B$5+A1", -1, -1)).toBe("Sheet1!$B$5+XFD1048576");
  });

  it("a relative name evaluates at its host: Above = Sheet1!A1048576 used in B5 is B4", () => {
    const book = model([sheet("s1", "Sheet1", { rows: [row(3, [cell(1, 7)]), row(4, [cell(1, null, "Above*2")])] })], [{ name: "Above", formula: "Sheet1!A1048576" }]);
    recalculate(book, [], { full: true, reason: "manual" });
    expect(valueAt(book, 0, 4, 1)).toBe(14);
  });

  it("a structural edit leaves a relative name's offsets alone (it still means the cell above)", () => {
    const book = model([sheet("s1", "Sheet1")], [{ name: "Above", formula: "Sheet1!A1048576" }]);
    applyStructuralEdit(book, { kind: "insertRows", sheetId: "s1", at: 2, count: 1 });
    expect(book.names[0]!.formula).toBe("Sheet1!A1048576");
  });
});

describe("names nest 64 deep", () => {
  /** Lvl_1 = Lvl_2 + … = Lvl_n = 5, A1 = Lvl_1. */
  function chain(depth: number, step: (next: string) => string = (next) => next): SpreadsheetModel {
    const names: DefinedName[] = [];
    for (let i = 1; i <= depth; i += 1) {
      names.push({ name: `Lvl_${i}`, formula: i === depth ? "5" : step(`Lvl_${i + 1}`) });
    }
    return model([sheet("s1", "Sheet1", { rows: [row(0, [cell(0, null, "Lvl_1")])] })], names);
  }

  it("a chain of 64 names evaluates; 65 is #NAME?", () => {
    const ok = chain(64);
    recalculate(ok, [], { full: true, reason: "manual" });
    expect(valueAt(ok, 0, 0, 0)).toBe(5);
    const deep = chain(65);
    recalculate(deep, [], { full: true, reason: "manual" });
    expect(valueAt(deep, 0, 0, 0)).toEqual({ error: "#NAME?" });
  });

  it("self and mutual references are #NAME?, and a doubling chain evaluates each name once", () => {
    const cycles = model([sheet("s1", "Sheet1", { rows: [row(0, [cell(0, null, "Foo"), cell(1, null, "Self")])] })], [
      { name: "Foo", formula: "Bar+Bar" },
      { name: "Bar", formula: "Foo+Foo" },
      { name: "Self", formula: "Self+Self" },
    ]);
    recalculate(cycles, [], { full: true, reason: "manual" });
    expect(valueAt(cycles, 0, 0, 0)).toEqual({ error: "#NAME?" });
    expect(valueAt(cycles, 0, 0, 1)).toEqual({ error: "#NAME?" });
    const doubling = chain(60, (next) => `${next}+${next}`);
    const started = performance.now();
    recalculate(doubling, [], { full: true, reason: "manual" });
    expect(performance.now() - started).toBeLessThan(2000);
    expect(valueAt(doubling, 0, 0, 0)).toBe(5 * 2 ** 59);
  });
});

describe("the names adjusters", () => {
  it("a row inserted above a name's range moves it (host: its scope sheet, else the edited sheet)", () => {
    const book = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")], [
      { name: "Margin", formula: "Sheet1!$B$2:$B$5" },
      { name: "Elsewhere", formula: "Sheet2!$B$2" },
      { name: "Local", scopeSheetId: "s1", formula: "$C$3" },
      { name: "Other", scopeSheetId: "s2", formula: "$C$3" },
      { name: "Raw", formula: "[1]Sheet1!$A$1", formulaUnparsed: true },
    ]);
    const names = book.names;
    applyStructuralEdit(book, { kind: "insertRows", sheetId: "s1", at: 0, count: 2 });
    expect(book.names.map((name) => name.formula)).toEqual(["Sheet1!$B$4:$B$7", "Sheet2!$B$2", "$C$5", "$C$3", "[1]Sheet1!$A$1"]);
    expect(book.names[1]).toBe(names[1]);
    expect(book.namesDirty).toBe(true);
    applyStructuralEdit(book, { kind: "deleteRows", sheetId: "s1", at: 0, count: 10 });
    expect(book.names[0]!.formula).toBe("Sheet1!#REF!");
  });

  it("an edit touching no name leaves the names array alone", () => {
    const book = model([sheet("s1", "Sheet1")], [{ name: "Margin", formula: "Sheet1!$B$2" }]);
    const names = book.names;
    applyStructuralEdit(book, { kind: "insertRows", sheetId: "s1", at: 5, count: 1 });
    expect(book.names).toBe(names);
    expect(book.namesDirty).toBeUndefined();
  });

  it("moveRange of a block containing a name's range moves it — across sheets with the new sheet", () => {
    const book = model([sheet("s1", "Sheet1"), sheet("s2", "Sheet2")], [{ name: "Block", formula: "Sheet1!$A$1:$B$2" }, { name: "Partly", formula: "Sheet1!$A$1:$C$3" }]);
    expect(moveRange(book, { fromSheetId: "s1", from: range(0, 0, 1, 1), toSheetId: "s1", toTop: 10, toLeft: 3 })).toEqual({ ok: true });
    expect(book.names.map((name) => name.formula)).toEqual(["Sheet1!$D$11:$E$12", "Sheet1!$A$1:$C$3"]);
    moveRange(book, { fromSheetId: "s1", from: range(10, 3, 11, 4), toSheetId: "s2", toTop: 0, toLeft: 0 });
    expect(book.names[0]!.formula).toBe("Sheet2!$A$1:$B$2");
  });
});

describe("SPEC-121's sheet commands reach names and tables", () => {
  function book(): SpreadsheetModel {
    return model(
      [
        sheet("s1", "Sheet1", {
          rows: [row(0, [cell(0, "Price"), cell(1, "Qty")]), row(1, [cell(0, 2), cell(1, 3)]), row(2, [cell(2, 0, "SUM(Sales[Price])")])],
          tables: [table("Sales", range(0, 0, 1, 1), { columns: [{ name: "Price", calculatedFormula: "Sheet2!$A$1*2" }, { name: "Qty", totalsFormula: "SUM(Sales[Qty])" }], xlsxId: 4, partName: "xl/tables/table1.xml" })],
        }),
        sheet("s2", "Sheet2", { rows: [row(0, [cell(0, 1), cell(1, 0, "SUM(Sales[Qty])")])] }),
      ],
      [
        { name: "Margin", formula: "Sheet2!$A$1" },
        { name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Sheet1!$A$1:$B$2" },
        { name: "Total", scopeSheetId: "s1", formula: "SUM(Sales[Price])" },
      ],
    );
  }

  it("rename rewrites every name's formula and every table's calculated and totals formulas", () => {
    const m = book();
    renameSheet(m, 1, "Data");
    expect(m.names[0]!.formula).toBe("Data!$A$1");
    expect(m.namesDirty).toBe(true);
    const sales = m.sheets[0]!.tables![0]!;
    expect(sales.columns[0]!.calculatedFormula).toBe("Data!$A$1*2");
    expect(sales.dirty).toBe(true);
  });

  it("delete drops the sheet's local names and tables, their references #REF!", () => {
    const m = book();
    deleteSheet(m, 0);
    expect(m.names.map((name) => name.name)).toEqual(["Margin"]);
    expect(getCell(m.sheets[0]!, 0, 1)?.formula).toBe("SUM(#REF!)");
    const other = book();
    deleteSheet(other, 1);
    expect(other.names[0]!.formula).toBe("#REF!");
    expect(other.sheets[0]!.tables![0]!.columns[0]!.calculatedFormula).toBe("#REF!*2");
  });

  it("duplicate copies tables (fresh id and xlsxId, <name>_2, no partName, dirty) and local names with the copy's scope", () => {
    const m = book();
    duplicateSheet(m, 0);
    const copy = m.sheets[1]!;
    const [copied] = copy.tables!;
    expect(copy.name).toBe("Sheet1 (2)");
    expect(copied).toMatchObject({ name: "Sales_2", xlsxId: 5, dirty: true });
    expect(copied!.id).not.toBe("id-Sales");
    expect(copied!.partName).toBeUndefined();
    expect(copied!.columns[1]!.totalsFormula).toBe("SUM(Sales_2[Qty])");
    expect(getCell(copy, 2, 2)?.formula).toBe("SUM(Sales_2[Price])");
    expect(getCell(m.sheets[0]!, 2, 2)?.formula).toBe("SUM(Sales[Price])");
    expect(copy.dirty).toBe(true);
    expect(m.names.slice(3)).toEqual([
      { name: "_xlnm.Print_Area", scopeSheetId: copy.id, formula: "'Sheet1 (2)'!$A$1:$B$2" },
      { name: "Total", scopeSheetId: copy.id, formula: "SUM(Sales_2[Price])" },
    ]);
    expect(m.namesDirty).toBe(true);
  });
});
