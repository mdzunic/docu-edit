// SPEC-133 round trip through .xlsx (xlsx/definedNames.ts, xlsx/tableXml.ts,
// xlsx/writer.ts): spec133-names-tables.xlsx untouched writes workbook.xml,
// both table parts and every other part byte-identical, and a save after an
// unrelated cell edit keeps `<definedNames>` and the table parts so; workbook
// and sheet names defined, edited and deleted survive save → reopen
// (equalStructural), untouched names keeping their bytes and new ones
// appended with `localSheetId` from their sheet's index, which moving a sheet
// recomputes and renaming one rewrites into the names' text; a table created,
// renamed, re-totalled, deleted or converted writes (or drops) its part,
// relationship, Override and `<tableParts>`, a regenerated part keeping the
// fixture's unmodelled XML; a duplicated sheet's table goes to its
// closure-copied part as Sales_2; a structural edit moves names and tables
// through the model alone (AC-22) and the writer writes the moved ones. Every
// reopened model writes byte-identical again. Opened and saved as the app
// does (readSpreadsheetBytes, writeXlsx).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { decodedPartBytes } from "../core/packageParts";
import { DEFAULT_THEME } from "../core/theme";
import { unzip } from "../core/zip";
import { equalStructural } from "../equality";
import { getCell, parseA1, parseRange } from "./address";
import { clearContents, setCellEntry } from "./cells";
import { buildSpec133NamesTablesXlsxBytes, SPEC133_NAMES_TABLES_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import type { DefinedName, SheetTable, SpreadsheetModel } from "./model";
import { defineName, deleteName, updateName } from "./names";
import { readSpreadsheetBytes } from "./open";
import { duplicateSheet, moveSheet, renameSheet } from "./sheets";
import { applyStructuralEdit } from "./structure";
import { convertTableToRange, createTable, renameTable, setTotalsRow } from "./tableCommands";
import { writeXlsx } from "./xlsx/writer";

const X = SPEC133_NAMES_TABLES_XLSX;
const { Sales, Table1 } = X.tables;
const TABLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const TABLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const WORKBOOK = "xl/workbook.xml";
const CALC_CHAIN = "xl/calcChain.xml";
/** A fixed instant for a rewritten docProps/core.xml — never "today". */
const MODIFIED = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
const decoder = new TextDecoder();
const encoder = new TextEncoder();

type Parts = Map<string, Uint8Array>;

const readXlsx = (bytes: Uint8Array): Promise<SpreadsheetModel> => readSpreadsheetBytes(bytes, "xlsx", "names");
const save = (model: SpreadsheetModel): Promise<Uint8Array> => writeXlsx(model, { modified: MODIFIED });

async function open(): Promise<{ bytes: Uint8Array; source: Parts; model: SpreadsheetModel }> {
  const bytes = await buildSpec133NamesTablesXlsxBytes();
  return { bytes, source: await unzip(bytes), model: await readXlsx(bytes) };
}

function text(parts: Parts, name: string): string {
  expect(parts.has(name), name).toBe(true);
  return decoder.decode(parts.get(name));
}

/** Whether a part is byte-identical in two packages. */
function samePart(a: Parts, b: Parts, name: string): boolean {
  expect(a.has(name) && b.has(name), name).toBe(true);
  return equalBytes(a.get(name)!, b.get(name)!);
}

/** Each part of a written package that differs from `other`'s. */
async function differingParts(a: Uint8Array, b: Uint8Array): Promise<string[]> {
  const left = await unzip(a);
  const right = await unzip(b);
  return [...new Set([...left.keys(), ...right.keys()])].filter((name) => !left.has(name) || !right.has(name) || !equalBytes(left.get(name)!, right.get(name)!));
}

/** The workbook part's `<definedNames>…</definedNames>`, verbatim. */
function definedNamesBlock(workbookXml: string): string {
  const start = workbookXml.indexOf("<definedNames>");
  const close = "</definedNames>";
  expect(start).toBeGreaterThanOrEqual(0);
  return workbookXml.slice(start, workbookXml.indexOf(close, start) + close.length);
}

/** A name's source `<definedName>` element, verbatim (the first of that name). */
function sourceElement(workbookXml: string, name: string): string {
  const start = workbookXml.indexOf(`<definedName name="${name}"`);
  const close = "</definedName>";
  expect(start, name).toBeGreaterThanOrEqual(0);
  return workbookXml.slice(start, workbookXml.indexOf(close, start) + close.length);
}

/** A model's names in model order — the order `<definedNames>` is written
 * in — each scope by its sheet's name, so two reads compare. */
function namesWithScope(model: SpreadsheetModel): unknown[] {
  return model.names.map(({ scopeSheetId, ...name }) => ({ ...name, scope: model.sheets.find((sheet) => sheet.id === scopeSheetId)?.name ?? null }));
}

/** A model's tables sheet by sheet, without what a save or a read mints —
 * `id`, `dirty`, `partName`, `xlsxId` (asserted on their own) — and with the
 * filter as its presence: a reader keeps the part's `<autoFilter>` verbatim
 * where a command holds `{}` or the source's unpatched text. */
function tablesBySheet(model: SpreadsheetModel): unknown[] {
  return model.sheets.map((sheet) => ({
    sheet: sheet.name,
    tables: (sheet.tables ?? []).map(({ id: _id, dirty: _dirty, partName: _part, xlsxId: _xlsxId, autoFilter, ...table }) => ({ ...table, filterButton: autoFilter !== undefined })),
  }));
}

function named(model: SpreadsheetModel, name: string, scope?: string): DefinedName {
  const scopeId = scope === undefined ? undefined : model.sheets.find((sheet) => sheet.name === scope)!.id;
  const found = model.names.find((each) => each.name === name && each.scopeSheetId === scopeId);
  expect(found, name).toBeDefined();
  return found!;
}

function tableNamed(model: SpreadsheetModel, name: string): { sheetIndex: number; table: SheetTable } {
  const sheetIndex = model.sheets.findIndex((sheet) => sheet.tables?.some((table) => table.name === name));
  expect(sheetIndex, name).toBeGreaterThanOrEqual(0);
  return { sheetIndex, table: model.sheets[sheetIndex]!.tables!.find((table) => table.name === name)! };
}

function sheetIndexOf(model: SpreadsheetModel, name: string): number {
  const index = model.sheets.findIndex((sheet) => sheet.name === name);
  expect(index, name).toBeGreaterThanOrEqual(0);
  return index;
}

function formulaAt(model: SpreadsheetModel, sheet: string, a1: string): string | undefined {
  const ref = parseA1(a1)!;
  return getCell(model.sheets[sheetIndexOf(model, sheet)]!, ref.row, ref.col)?.formula;
}

/** The Override a table part takes in [Content_Types].xml. */
const tableOverride = (partName: string): string => `<Override PartName="/${partName}" ContentType="${TABLE_CONTENT_TYPE}"/>`;

/** Reopens `written`, expects its names and tables equalStructural to
 * `model`'s, and a second write byte-identical to `written`. */
async function reopenAgrees(model: SpreadsheetModel, written: Uint8Array): Promise<SpreadsheetModel> {
  const reopened = await readXlsx(written);
  expect(reopened.sheets.map((sheet) => sheet.name)).toEqual(model.sheets.map((sheet) => sheet.name));
  expect(equalStructural(namesWithScope(reopened), namesWithScope(model))).toBe(true);
  expect(equalStructural(tablesBySheet(reopened), tablesBySheet(model))).toBe(true);
  expect(await differingParts(await save(reopened), written)).toEqual([]);
  return reopened;
}

describe("SPEC-133 round trip — spec133-names-tables.xlsx, untouched and clean", () => {
  it("an untouched workbook writes workbook.xml and both table parts byte-identical (and every other part)", async () => {
    const { bytes, source, model } = await open();
    expect(model.names.map((name) => name.name)).toEqual(X.definedNames.map((name) => name.name));
    expect(model.sheets.map((sheet) => (sheet.tables ?? []).map((table) => [table.name, table.partName, table.xlsxId]))).toEqual([
      [[Sales.name, Sales.partName, Sales.id]],
      [[Table1.name, Table1.partName, Table1.id]],
      [],
    ]);
    const written = await save(model);
    const parts = await unzip(written);
    for (const name of [WORKBOOK, Sales.partName, Table1.partName]) {
      expect(samePart(parts, source, name), name).toBe(true);
    }
    expect(await differingParts(written, bytes)).toEqual([]);
  });

  it("a save after an unrelated cell edit on another sheet keeps <definedNames> and both table parts byte-identical", async () => {
    const { source, model } = await open();
    const hidden = sheetIndexOf(model, "Hidden");
    setCellEntry(model, hidden, { row: 0, col: 1 }, { value: 43 });
    model.sheets[hidden] = { ...model.sheets[hidden]!, dirty: true };
    const written = await save(model);
    const parts = await unzip(written);
    expect(await packageIntegrity(written)).toEqual([]);
    for (const name of [Sales.partName, Table1.partName, X.sheets.Data.partName, X.sheets.Lists.partName, X.sheets.Data.rels, X.sheets.Lists.rels]) {
      expect(samePart(parts, source, name), name).toBe(true);
    }
    // The names are not rebuilt: the block keeps its bytes. The one change to
    // the part is SPEC-121's substitution-path `workbookView@activeTab`.
    const sourceWorkbook = text(source, WORKBOOK);
    expect(definedNamesBlock(text(parts, WORKBOOK))).toBe(definedNamesBlock(sourceWorkbook));
    const expected = sourceWorkbook.replace(/(<workbookView\b[^>]*?)\/>/, `$1 activeTab="${X.activeSheetIndex}"/>`);
    expect(expected).not.toBe(sourceWorkbook);
    expect(equalBytes(parts.get(WORKBOOK)!, encoder.encode(expected))).toBe(true);
    expect(getCell((await readXlsx(written)).sheets[hidden]!, 0, 1)?.value).toBe(43);
  });
});

describe("SPEC-133 round trip — names", () => {
  it("workbook and sheet names defined, edited and deleted → save → reopen: the names equalStructural; untouched names keep their bytes; a second write is byte-identical", async () => {
    const { source, model } = await open();
    const sourceWorkbook = text(source, WORKBOOK);
    const hidden = model.sheets[sheetIndexOf(model, "Hidden")]!;
    // Workbook scope: a new range, an edited constant, a deleted range.
    defineName(model, { name: "Margin", formula: "Data!$B$2:$B$5" });
    const taxRate = named(model, "TaxRate");
    updateName(model, taxRate, { ...taxRate, formula: "0.25" });
    deleteName(model, named(model, "Rates"));
    // Sheet scope: a renamed range (its comment kept), a deleted relative name, a new expression on Hidden.
    const local = named(model, "Local", "Lists");
    updateName(model, local, { ...local, name: "Codes" });
    deleteName(model, named(model, "Above", "Data"));
    defineName(model, { name: "Half", scopeSheetId: hidden.id, formula: "Hidden!$B$1/2" });
    expect(model.namesDirty).toBe(true);

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    const workbook = text(parts, WORKBOOK);
    expect(definedNamesBlock(workbook)).toBe(
      `<definedNames>` +
        `<definedName name="_xlfn.IFS" hidden="1">#NAME?</definedName>` +
        `<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Lists!$D$1:$E$4</definedName>` +
        `<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$E$6</definedName>` +
        `<definedName name="Helper" hidden="1">Data!$Z$1</definedName>` +
        `<definedName name="Codes" comment="Codes on Lists" localSheetId="1">Lists!$A$2:$A$4</definedName>` +
        `<definedName name="TaxRate">0.25</definedName>` +
        // New names last, in the order defined; a scoped one's localSheetId is its sheet's index.
        `<definedName name="Margin">Data!$B$2:$B$5</definedName>` +
        `<definedName name="Half" localSheetId="2">Hidden!$B$1/2</definedName>` +
        `</definedNames>`,
    );
    for (const name of ["_xlfn.IFS", "_xlnm._FilterDatabase", "_xlnm.Print_Area", "Helper"]) {
      expect(workbook, name).toContain(sourceElement(sourceWorkbook, name));
    }
    // Names are not table parts: both stay as they were.
    expect(samePart(parts, source, Sales.partName)).toBe(true);
    expect(samePart(parts, source, Table1.partName)).toBe(true);

    const reopened = await reopenAgrees(model, first);
    expect(named(reopened, "Half", "Hidden").formula).toBe("Hidden!$B$1/2");
    expect(named(reopened, "Codes", "Lists").comment).toBe("Codes on Lists");
    expect(reopened.names.some((name) => name.name === "Rates" || name.name === "Above" || name.name === "Local")).toBe(false);
  });

  it("moving a sheet recomputes localSheetId from the scope sheet's current index; renaming one rewrites the names' text", async () => {
    const { source, model } = await open();
    moveSheet(model, sheetIndexOf(model, "Lists"), 0);
    renameSheet(model, sheetIndexOf(model, "Data"), "Facts");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Lists", "Facts", "Hidden"]);

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    expect(definedNamesBlock(text(parts, WORKBOOK))).toBe(
      `<definedNames>` +
        `<definedName name="_xlfn.IFS" hidden="1">#NAME?</definedName>` +
        `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Lists!$D$1:$E$4</definedName>` +
        `<definedName name="_xlnm.Print_Area" localSheetId="1">Facts!$A$1:$E$6</definedName>` +
        `<definedName name="Above" localSheetId="1">Facts!A1048576</definedName>` +
        `<definedName name="Helper" hidden="1">Facts!$Z$1</definedName>` +
        `<definedName name="Local" comment="Codes on Lists" localSheetId="0">Lists!$A$2:$A$4</definedName>` +
        `<definedName name="Rates">Lists!$B$2:$B$4</definedName>` +
        `<definedName name="TaxRate">0.2</definedName>` +
        `</definedNames>`,
    );
    // The tables name no sheet: their parts are untouched.
    expect(samePart(parts, source, Sales.partName)).toBe(true);
    expect(samePart(parts, source, Table1.partName)).toBe(true);

    const reopened = await reopenAgrees(model, first);
    expect(named(reopened, "Above", "Facts").formula).toBe("Facts!A1048576");
    expect(named(reopened, "Local", "Lists").formula).toBe("Lists!$A$2:$A$4");
  });
});

describe("SPEC-133 round trip — tables", () => {
  it("a table created through createTable → save: the next free xl/tables/table{n}.xml with a workbook-unique id, the sheet relationship, the Override and <tableParts> → reopen equalStructural", async () => {
    const { source, model } = await open();
    const data = sheetIndexOf(model, "Data");
    // A fresh block right of column G's formulas.
    const block: [string, string | number][] = [
      ["I1", "Region"],
      ["J1", "Units"],
      ["I2", "North"],
      ["J2", 5],
      ["I3", "South"],
      ["J3", 7],
    ];
    for (const [a1, value] of block) {
      setCellEntry(model, data, parseA1(a1)!, { value });
    }
    const created = createTable(model, model.sheets[data]!.id, parseRange("I1:J3")!, true);
    expect(created.ok).toBe(true);
    expect(tableNamed(model, "Table2").table).toMatchObject({ ref: parseRange("I1:J3"), headerRow: true, columns: [{ name: "Region" }, { name: "Units" }] });

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    const added = "xl/tables/table3.xml";
    expect([...parts.keys()].filter((name) => !source.has(name))).toEqual([added]);
    expect(text(parts, added)).toContain(` id="3" name="Table2" displayName="Table2" ref="I1:J3"`);
    expect(text(parts, "[Content_Types].xml")).toContain(tableOverride(added));
    const rels = text(parts, X.sheets.Data.rels);
    expect(rels).toContain(`<Relationship Id="${Sales.rId}" Type="${TABLE_REL_TYPE}" Target="../tables/table1.xml"/>`);
    expect(rels).toContain(`<Relationship Id="rId2" Type="${TABLE_REL_TYPE}" Target="../tables/table3.xml"/>`);
    expect(text(parts, X.sheets.Data.partName)).toMatch(/<tableParts count="2"><tablePart r:id="rId1"\/><tablePart r:id="rId2"\/><\/tableParts><\/worksheet>$/);
    // The fixture's tables are clean: their parts are untouched.
    expect(samePart(parts, source, Sales.partName)).toBe(true);
    expect(samePart(parts, source, Table1.partName)).toBe(true);

    const reopened = await reopenAgrees(model, first);
    expect(tableNamed(reopened, "Table2").table).toMatchObject({ partName: added, xlsxId: 3 });
  });

  it("renaming a table and toggling its totals row regenerate its part keeping the fixture's unmodelled XML (xr:uid, xr3:uid, dataDxfId) → reopen equalStructural", async () => {
    const { source, model } = await open();
    const dataId = model.sheets[sheetIndexOf(model, "Data")]!.id;
    const sales = tableNamed(model, "Sales").table;
    expect(renameTable(model, sales.id, "Orders")).toEqual({ ok: true });
    expect(setTotalsRow(model, dataId, sales.id, false)).toEqual({ ok: true });
    expect(formulaAt(model, "Data", "G2")).toBe("SUM(Orders[Amount])");
    // Excel's `Sales['#Units]` in the totals cell reads back as the count this module writes, not "custom".
    const units = tableNamed(model, "Orders").table.columns.find((column) => column.name === "#Units")!;
    expect(units.totalsFunction).toBe("count");
    expect(units.totalsFormula).toBeUndefined();

    /** The regenerated part keeps every uid and the Amount column's dataDxfId and calculated formula. */
    const keepsUnmodelled = (xml: string): void => {
      expect(xml).toContain(` id="${Sales.id}" xr:uid="${Sales.uid}" name="Orders" displayName="Orders"`);
      expect(xml).toContain(`<autoFilter ref="A1:E5" xr:uid="${Sales.autoFilter.uid}"/>`);
      for (const column of Sales.columns) {
        expect(xml).toContain(`<tableColumn id="${column.id}" xr3:uid="${column.uid}" name="${column.name}"`);
      }
      expect(xml).toContain(` name="Amount" totalsRowFunction="sum" dataDxfId="${X.dxf.id}"><calculatedColumnFormula>Orders[[#This Row],[Price]]*Orders[[#This Row],[Qty]]</calculatedColumnFormula>`);
      expect(xml).toContain(`mc:Ignorable="xr xr3"`);
    };

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    const off = text(parts, Sales.partName);
    keepsUnmodelled(off);
    expect(off).toContain(` displayName="Orders" ref="A1:E5">`);
    expect(off).not.toContain("totalsRowCount");
    expect(samePart(parts, source, Table1.partName)).toBe(true);
    expect(text(parts, "[Content_Types].xml")).toContain(tableOverride(Sales.partName));
    expect(samePart(parts, source, X.sheets.Data.rels)).toBe(true);

    const reopened = await reopenAgrees(model, first);
    expect(tableNamed(reopened, "Orders").table).toMatchObject({ ref: parseRange("A1:E5"), totalsRow: false, xlsxId: Sales.id, partName: Sales.partName });

    // And back on: the totals row returns to row 6 in the same part.
    const orders = tableNamed(reopened, "Orders").table;
    expect(setTotalsRow(reopened, reopened.sheets[sheetIndexOf(reopened, "Data")]!.id, orders.id, true)).toEqual({ ok: true });
    const second = await save(reopened);
    expect(await packageIntegrity(second)).toEqual([]);
    const on = text(await unzip(second), Sales.partName);
    keepsUnmodelled(on);
    expect(on).toContain(` displayName="Orders" ref="A1:E6" totalsRowCount="1">`);
    const again = await reopenAgrees(reopened, second);
    expect(tableNamed(again, "Orders").table).toMatchObject({ ref: parseRange(Sales.ref), totalsRow: true });
  });

  it("a deleted table and a converted one: their parts, relationships, Overrides and <tableParts> leave the package", async () => {
    const { source, model } = await open();
    // Clearing every cell of Sales deletes it; Table1 is converted to a range.
    clearContents(model, sheetIndexOf(model, "Data"), [parseRange(Sales.ref)!]);
    expect(model.sheets[sheetIndexOf(model, "Data")]!.tables).toEqual([]);
    const lists = model.sheets[sheetIndexOf(model, "Lists")]!;
    expect(convertTableToRange(model, lists.id, tableNamed(model, "Table1").table.id, model.theme ?? DEFAULT_THEME)).toEqual({ ok: true });

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    expect([...parts.keys()].filter((name) => name.startsWith("xl/tables/"))).toEqual([]);
    // Nothing else of the package goes (calcChain is SPEC-122's to drop).
    expect([...source.keys()].filter((name) => !parts.has(name))).toEqual([Sales.partName, Table1.partName, CALC_CHAIN]);
    const contentTypes = text(parts, "[Content_Types].xml");
    expect(contentTypes).not.toContain(TABLE_CONTENT_TYPE);
    for (const sheet of [X.sheets.Data, X.sheets.Lists]) {
      expect(parts.has(sheet.rels) ? text(parts, sheet.rels) : "", sheet.rels).not.toContain(TABLE_REL_TYPE);
      expect(text(parts, sheet.partName), sheet.partName).not.toContain("tableParts");
    }

    const reopened = await reopenAgrees(model, first);
    expect(reopened.sheets.every((sheet) => (sheet.tables ?? []).length === 0)).toBe(true);
  });

  it("duplicating the sheet holding Sales: the copy's table, Sales_2 with a workbook-unique id, goes to the closure-copied part, nothing of the package dropped; reopen sees both tables and the copy's formulas read Sales_2[...]", async () => {
    const { source, model } = await open();
    duplicateSheet(model, sheetIndexOf(model, "Data"));
    const copyIndex = sheetIndexOf(model, "Data (2)");
    const copy = model.sheets[copyIndex]!;
    const copied = copy.tables![0]!;
    expect(copied).toMatchObject({ name: "Sales_2", dirty: true });
    // sheets.ts gives the copy's table the part SPEC-121's closure copy made.
    const copiedPart = copied.partName!;
    expect(copiedPart).toMatch(/^xl\/tables\/[^/]+\.xml$/);
    expect(source.has(copiedPart)).toBe(false);
    expect(decodedPartBytes(model.opaque).has(copiedPart)).toBe(true);

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    // Every source part but calcChain stays; the copy brings its sheet, its rels and its table.
    expect([...source.keys()].filter((name) => !parts.has(name))).toEqual([CALC_CHAIN]);
    expect([...parts.keys()].filter((name) => !source.has(name)).toSorted()).toEqual([copiedPart, copy.partName!, `xl/worksheets/_rels/${copy.partName!.split("/").pop()!}.rels`].toSorted());
    const copyXml = text(parts, copiedPart);
    expect(copyXml).toContain(` id="3" `);
    expect(copyXml).toContain(` name="Sales_2" displayName="Sales_2" ref="${Sales.ref}" totalsRowCount="1">`);
    expect(copyXml).toContain(`<calculatedColumnFormula>Sales_2[[#This Row],[Price]]*Sales_2[[#This Row],[Qty]]</calculatedColumnFormula>`);
    expect(text(parts, "[Content_Types].xml")).toContain(tableOverride(copiedPart));
    expect(text(parts, `xl/worksheets/_rels/${copy.partName!.split("/").pop()!}.rels`)).toContain(`Target="../tables/${copiedPart.split("/").pop()!}"`);
    expect(text(parts, copy.partName!)).toContain(`<tableParts count="1">`);
    // The origin's tables stay as they were.
    expect(samePart(parts, source, Sales.partName)).toBe(true);
    expect(samePart(parts, source, Table1.partName)).toBe(true);

    const reopened = await reopenAgrees(model, first);
    expect(reopened.sheets.flatMap((sheet) => (sheet.tables ?? []).map((table) => [sheet.name, table.name, table.xlsxId, table.partName]))).toEqual([
      ["Data", "Sales", Sales.id, Sales.partName],
      ["Data (2)", "Sales_2", 3, copiedPart],
      ["Lists", "Table1", Table1.id, Table1.partName],
    ]);
    expect(formulaAt(reopened, "Data (2)", "G2")).toBe("SUM(Sales_2[Amount])");
    expect(formulaAt(reopened, "Data (2)", "G3")).toBe("Sales_2[[#Totals],[Amount]]");
    expect(formulaAt(reopened, "Data (2)", "D2")).toBe("Sales_2[@Price]*Sales_2[@Qty]");
    expect(formulaAt(reopened, "Data (2)", "E6")).toBe("SUBTOTAL(103,Sales_2['#Units])");
    expect(formulaAt(reopened, "Data", "G2")).toBe(X.formulas.G2.formula);
    // The copy's local names are its own.
    expect(named(reopened, "Above", "Data (2)").formula).toBe("'Data (2)'!A1048576");
    expect(named(reopened, "_xlnm.Print_Area", "Data (2)").formula).toBe("'Data (2)'!$A$1:$E$6");
  });
});

describe("SPEC-133 round trip — a structural edit", () => {
  it("a row insert above the table and the names moves them through the model alone — the retained parts keep their bytes (AC-22) — and the save writes the moved names and table ref; reopen agrees", async () => {
    const { source, model } = await open();
    const data = model.sheets[sheetIndexOf(model, "Data")]!;
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: data.id, at: 0, count: 2 })).toEqual({ ok: true });
    expect(named(model, "_xlnm.Print_Area", "Data").formula).toBe("Data!$A$3:$E$8");
    expect(named(model, "Helper").formula).toBe("Data!$Z$3");
    expect(named(model, "Local", "Lists").formula).toBe("Lists!$A$2:$A$4");
    expect(tableNamed(model, "Sales").table).toMatchObject({ ref: parseRange("A3:E8"), dirty: true });
    // SPEC-124's retained rows for definedName text and table@ref/autoFilter@ref are retired.
    const retained = decodedPartBytes(model.opaque);
    for (const name of [WORKBOOK, Sales.partName, Table1.partName]) {
      expect(equalBytes(retained.get(name)!, source.get(name)!), name).toBe(true);
    }

    const first = await save(model);
    const parts = await unzip(first);
    expect(await packageIntegrity(first)).toEqual([]);
    const workbook = text(parts, WORKBOOK);
    expect(workbook).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$3:$E$8</definedName>`);
    expect(workbook).toContain(`<definedName name="Helper" hidden="1">Data!$Z$3</definedName>`);
    const sourceWorkbook = text(source, WORKBOOK);
    for (const name of ["_xlfn.IFS", "_xlnm._FilterDatabase", "Local", "Rates", "TaxRate"]) {
      expect(workbook, name).toContain(sourceElement(sourceWorkbook, name));
    }
    const table = text(parts, Sales.partName);
    expect(table).toContain(` name="Sales" displayName="Sales" ref="A3:E8" totalsRowCount="1"><autoFilter ref="A3:E7" xr:uid="${Sales.autoFilter.uid}"/>`);
    expect(samePart(parts, source, Table1.partName)).toBe(true);

    const reopened = await reopenAgrees(model, first);
    expect(tableNamed(reopened, "Sales").table.ref).toEqual(parseRange("A3:E8"));
    expect(named(reopened, "_xlnm.Print_Area", "Data").formula).toBe("Data!$A$3:$E$8");
    expect(formulaAt(reopened, "Data", "G4")).toBe(X.formulas.G2.formula);
  });
});
