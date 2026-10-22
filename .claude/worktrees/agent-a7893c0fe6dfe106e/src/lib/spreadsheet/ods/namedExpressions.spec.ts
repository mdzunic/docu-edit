// SPEC-133 "Names on the wire" (ods) and "Tables in ODS"
// (ods/namedExpressions.ts): reading the spreadsheet's and the tables' named
// ranges and expressions rebased to A1, writing the model's names back
// (named range vs expression, base cells, where the element goes, verbatim
// when unchanged), and tables written as database ranges with their look
// baked in and their structured references as ranges.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { buildSpec133NamesOdsBytes, SPEC133_NAMES_ODS } from "../fixtures/build";
import { MAX_COLS, MAX_ROWS, type DefinedName, type Sheet, type SheetTable, type SpreadsheetModel } from "../model";
import { odsNameXml, odsNamedRangeAddress, odsSuffixWithNames, odsWithTableDatabaseRanges } from "./namedExpressions";
import { parseOds } from "./reader";
import { writeOds } from "./writer";

const OFFICE = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"`;
const MANIFEST = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
const encoder = new TextEncoder();
const decoder = new TextDecoder();

const cell = (value: number): string => `<table:table-cell office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`;
const table = (name: string, trailing = ""): string => `<table:table table:name="${name}"><table:table-column/><table:table-row>${cell(1)}</table:table-row>${trailing}</table:table>`;

function content(tables: string, suffix: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${OFFICE} office:version="1.2"><office:body><office:spreadsheet>${tables}${suffix}</office:spreadsheet></office:body></office:document-content>`;
}

function odsBytes(xml: string): Promise<Uint8Array> {
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    { name: "META-INF/manifest.xml", data: encoder.encode(MANIFEST) },
    { name: "content.xml", data: encoder.encode(xml) },
  ]);
}

/** A model opened from bytes, as open.ts assembles one. */
async function open(bytes: Uint8Array): Promise<SpreadsheetModel> {
  const parsed = await parseOds(bytes);
  return {
    kind: "spreadsheet",
    id: "model",
    title: "Names",
    format: "ods",
    readerVersion: 0,
    sheets: parsed.sheets,
    activeSheetIndex: parsed.activeSheetIndex,
    dateSystem: parsed.dateSystem,
    opaque: parsed.opaque,
    styles: parsed.styles,
    defaultStyleId: 0,
    names: parsed.names ?? [],
    dxfs: parsed.dxfs,
    ...(parsed.odsContentEnvelope ? { odsContentEnvelope: parsed.odsContentEnvelope } : {}),
  };
}

async function contentOf(bytes: Uint8Array): Promise<string> {
  return decoder.decode((await unzip(bytes)).get("content.xml")!);
}

/** The text of the `table:table` named `name` in content.xml. */
function tableText(xml: string, name: string): string {
  const start = xml.indexOf(`<table:table table:name="${name}"`);
  return xml.slice(start, xml.indexOf("</table:table>", start) + "</table:table>".length);
}

const BASE_NAMES = `<table:named-expressions>
  <table:named-range table:name="Below" table:base-cell-address="$Data.$B$3" table:cell-range-address="$Data.B4" table:range-usable-as="repeat-row"/>
  <table:named-range table:name="Wrapped" table:base-cell-address="$Data.$C$3" table:cell-range-address="$Data.A1:.$B2"/>
  <table:named-expression table:name="Twice" table:base-cell-address="$Data.$D$5" table:expression="of:=[.C5]*2+[$Data.$A$1]"/>
  <table:named-range table:name="Gone" table:base-cell-address="$Data.$A$1" table:cell-range-address="$#REF!.$A$1"/>
  <table:named-expression table:name="Odd" table:base-cell-address="$Data.$A$1" table:expression="zz:=1"/>
</table:named-expressions>`;
const LOCAL_NAMES = `<table:named-expressions><table:named-range table:name="Own" table:base-cell-address="$Data.$B$2" table:cell-range-address="$Data.B2:.B3"/></table:named-expressions>`;

describe("SPEC-133 ods names: reading", () => {
  it("reads the workbook's names, then each table's, rebased from their base cells to A1 (wrapping)", async () => {
    const model = await open(await odsBytes(content(table("Data", LOCAL_NAMES) + table("Other"), BASE_NAMES)));
    const data = model.sheets[0]!;
    expect(model.names).toEqual<DefinedName[]>([
      { name: "Below", formula: "Data!A2", sourceAttrs: ` table:range-usable-as="repeat-row"` },
      // B3-relative A1 → row −2, col −2: wraps to the grid's far edge; the absolute $B keeps its column.
      { name: "Wrapped", formula: `Data!$B${MAX_ROWS - 1}:XFC${MAX_ROWS}` },
      { name: "Twice", formula: `${"XFD"}1*2+Data!$A$1` },
      { name: "Gone", formula: "$#REF!.$A$1", formulaUnparsed: true, sourceAttrs: ` table:base-cell-address="$Data.$A$1" table:cell-range-address="$#REF!.$A$1"` },
      { name: "Odd", formula: "zz:=1", formulaUnparsed: true, sourceAttrs: ` table:base-cell-address="$Data.$A$1" table:expression="zz:=1"` },
      { name: "Own", scopeSheetId: data.id, formula: "Data!A1:A2" },
    ]);
    expect(MAX_COLS).toBe(16384);
  });

  it("reads LibreOffice's fixture — and never makes a table of a database range", async () => {
    const model = await open(await buildSpec133NamesOdsBytes());
    const sheet1 = model.sheets.find((sheet) => sheet.name === "Sheet1")!;
    const expected = [
      ...SPEC133_NAMES_ODS.globalNames.map((name) => ({ name: name.name, formula: name.formula })),
      ...SPEC133_NAMES_ODS.localNames.map((name) => ({ name: name.name, formula: name.formula, scopeSheetId: sheet1.id })),
    ];
    expect(model.names.map(({ name, formula, scopeSheetId }) => ({ name, formula, ...(scopeSheetId === undefined ? {} : { scopeSheetId }) }))).toEqual(expected);
    expect(model.names.find((name) => name.name === "Report")?.sourceAttrs).toBe(` table:range-usable-as="print-range filter"`);
    expect(model.sheets.every((sheet) => sheet.tables === undefined)).toBe(true);
  });

  it("re-emits an untouched workbook's content.xml byte-identical", async () => {
    const bytes = await buildSpec133NamesOdsBytes();
    const written = await writeOds(await open(bytes));
    expect(equalBytes((await unzip(written)).get("content.xml")!, (await unzip(bytes)).get("content.xml")!)).toBe(true);
  });
});

/** A new one-sheet workbook (no retained package: written from scratch). */
function blank(): SpreadsheetModel {
  const sheet: Sheet = { id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], conditionalFormats: [], validations: [], dirty: true };
  return { kind: "spreadsheet", id: "new", title: "New", format: "ods", readerVersion: 0, sheets: [sheet], activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0, names: [], dxfs: [] };
}

describe("SPEC-133 ods names: writing", () => {
  const model = blank();
  const sheet = model.sheets[0]!;

  it("writes a sheet-qualified cell or range as a named range based at A1 of its scope (else the first) sheet", () => {
    expect(odsNameXml(model, { name: "R", formula: `${sheet.name}!$A$1:B3` })).toBe(
      `<table:named-range table:name="R" table:base-cell-address="$${sheet.name}.$A$1" table:cell-range-address="$${sheet.name}.$A$1:.B3"/>`,
    );
    expect(odsNameXml(model, { name: "P", formula: `${sheet.name}!$A$1`, sourceAttrs: ` table:range-usable-as="print-range"` })).toContain(`table:cell-range-address="$${sheet.name}.$A$1" table:range-usable-as="print-range"/>`);
    expect(odsNamedRangeAddress("Sheet1:Sheet3!B5")).toBe("$Sheet1.B5:$Sheet3.B5");
  });

  it("writes anything else — a constant, an unqualified or whole-column reference, a union — as a named expression", () => {
    expect(odsNameXml(model, { name: "Rate", formula: "0.15" })).toBe(`<table:named-expression table:name="Rate" table:base-cell-address="$${sheet.name}.$A$1" table:expression="of:=0.15"/>`);
    expect(odsNameXml(model, { name: "Near", formula: "B2*2", sourceAttrs: ` table:range-usable-as="filter"` })).toBe(
      `<table:named-expression table:name="Near" table:base-cell-address="$${sheet.name}.$A$1" table:expression="of:=[.B2]*2"/>`,
    );
    for (const formula of ["A1:B2", `${sheet.name}!$A:$A`, `(${sheet.name}!A1,${sheet.name}!C1)`]) {
      expect(odsNamedRangeAddress(formula)).toBeNull();
    }
  });

  it("puts the workbook's element right after the last table when the suffix has none, and replaces it otherwise", () => {
    const named: SpreadsheetModel = { ...model, names: [{ name: "Rate", formula: "0.15" }, { name: "_xlnm.Print_Area", scopeSheetId: sheet.id, formula: `${sheet.name}!$A$1` }] };
    const element = `<table:named-expressions>${odsNameXml(named, named.names[0]!)}</table:named-expressions>`;
    expect(odsSuffixWithNames("\n<table:database-ranges/></office:spreadsheet>", named, new Map()).suffix).toBe(`${element}\n<table:database-ranges/></office:spreadsheet>`);
    expect(odsSuffixWithNames(`<table:named-expressions><table:named-range table:name="Old" table:cell-range-address="$S.A1"/></table:named-expressions><x/>`, named, new Map()).suffix).toBe(`${element}<x/>`);
    // Nothing to write: the element goes.
    expect(odsSuffixWithNames(`<table:named-expressions><table:named-range table:name="Old" table:cell-range-address="$S.A1"/></table:named-expressions><x/>`, model, new Map()).suffix).toBe(`<x/>`);
  });

  it("writes a changed name of a clean workbook, the sheet's own names as the last child of their table, and reads them back", async () => {
    const source = await open(await odsBytes(content(table("Data", LOCAL_NAMES) + table("Other"), BASE_NAMES)));
    const data = source.sheets[0]!;
    const edited: SpreadsheetModel = {
      ...source,
      names: [...source.names.filter((name) => name.name !== "Twice"), { name: "Extra", scopeSheetId: data.id, formula: "Other!$A$1:$A$9" }],
      namesDirty: true,
    };
    const xml = await contentOf(await writeOds(edited));
    expect(tableText(xml, "Data").endsWith(
      // The scope changed, so its element is the model's: "Own" re-based at A1.
      `<table:named-expressions><table:named-range table:name="Own" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.A1:.A2"/>` +
        `<table:named-range table:name="Extra" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Other.$A$1:.$A$9"/></table:named-expressions></table:table>`,
    )).toBe(true);
    expect(tableText(xml, "Other")).toBe(table("Other"));
    expect(xml).not.toContain(`table:name="Twice"`);
    const reread = await open(await writeOds(edited));
    expect(reread.names.map(({ scopeSheetId: _scope, ...name }) => name)).toEqual(edited.names.map(({ scopeSheetId: _scope, ...name }) => name));
  });

  it("keeps the retained elements verbatim when a dirty sheet's names did not change", async () => {
    const source = await open(await odsBytes(content(table("Data", LOCAL_NAMES) + table("Other"), BASE_NAMES)));
    const dirty: SpreadsheetModel = { ...source, sheets: source.sheets.map((each, i) => (i === 0 ? { ...each, dirty: true as const } : each)) };
    const xml = await contentOf(await writeOds(dirty));
    expect(xml).toContain(`${LOCAL_NAMES}</table:table>`);
    expect(xml).toContain(BASE_NAMES);
  });

  it("leaves every names element as written for a model with no names to speak of — a record a pre-SPEC-133 reader stored, written through the substitution layout while its reader upgrade re-reads it", async () => {
    const source = await open(await odsBytes(content(table("Data", LOCAL_NAMES) + table("Other"), BASE_NAMES)));
    const { names: _names, ...rest } = source;
    const stored = { ...rest, sheets: source.sheets.map((each, i) => (i === 0 ? { ...each, dirty: true as const } : each)) } as SpreadsheetModel;
    expect(stored.names).toBeUndefined();
    const xml = await contentOf(await writeOds(stored));
    expect(xml).toContain(`${LOCAL_NAMES}</table:table>`);
    expect(tableText(xml, "Other")).toBe(table("Other"));
    expect(xml).toContain(BASE_NAMES);
  });

  it("writes new names from scratch: the workbook's after the tables, a sheet's last in its table", async () => {
    const book = blank();
    const first = book.sheets[0]!;
    book.names = [{ name: "Rate", formula: "0.15" }, { name: "Here", scopeSheetId: first.id, formula: `${first.name}!$B$2` }];
    const xml = await contentOf(await writeOds(book));
    expect(xml).toContain(`<table:named-expressions><table:named-range table:name="Here" table:base-cell-address="$${first.name}.$A$1" table:cell-range-address="$${first.name}.$B$2"/></table:named-expressions></table:table>`);
    expect(xml).toContain(`</table:table><table:named-expressions><table:named-expression table:name="Rate"`);
  });
});

describe("SPEC-133 ods tables", () => {
  const SALES: SheetTable = {
    id: "t1",
    name: "Sales",
    ref: { top: 0, left: 0, bottom: 3, right: 2 },
    headerRow: true,
    totalsRow: false,
    columns: [{ name: "Item" }, { name: "Amount" }, { name: "Calc" }],
    style: { name: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    autoFilter: {},
  };

  function withTable(): SpreadsheetModel {
    const book = blank();
    const sheet: Sheet = {
      ...book.sheets[0]!,
      rows: [
        { index: 0, cells: [{ col: 0, value: "Item" }, { col: 1, value: "Amount" }, { col: 2, value: "Calc" }] },
        { index: 1, cells: [{ col: 0, value: "a" }, { col: 1, value: 1 }] },
        { index: 2, cells: [{ col: 0, value: "b" }, { col: 1, value: 2 }] },
        { index: 3, cells: [{ col: 0, value: "c" }, { col: 1, value: 3 }, { col: 2, value: 3, formula: "Sales[@Amount]" }] },
        { index: 5, cells: [{ col: 1, value: 6, formula: "SUBTOTAL(109,Sales[Amount])" }, { col: 2, value: 9, formula: "SUM(Sales)" }] },
      ],
      tables: [SALES],
    };
    return { ...book, sheets: [sheet] };
  }

  it("writes each table as a database range named after it", () => {
    expect(odsWithTableDatabaseRanges("", [{ name: "My Sheet", tables: [SALES, { ...SALES, name: "Bare", headerRow: false, autoFilter: undefined }] }]).text).toBe(
      `<table:database-ranges><table:database-range table:name="Sales" table:target-range-address="'My Sheet'.A1:'My Sheet'.C4" table:contains-header="true" table:display-filter-buttons="true"/>` +
        `<table:database-range table:name="Bare" table:target-range-address="'My Sheet'.A1:'My Sheet'.C4" table:contains-header="false" table:display-filter-buttons="false"/></table:database-ranges>`,
    );
  });

  it("keeps the table's name, renaming another database range of that name", () => {
    const suffix = `<table:named-expressions/><table:database-ranges><table:database-range table:name="SALES" table:target-range-address="S.A1:S.B2"/><table:database-range table:name="Sales_2" table:target-range-address="S.A1:S.B2"/></table:database-ranges></office:spreadsheet>`;
    const { text } = odsWithTableDatabaseRanges(suffix, [{ name: "S", tables: [SALES] }]);
    expect(text).toContain(`<table:database-range table:name="SALES_3" table:target-range-address="S.A1:S.B2"/><table:database-range table:name="Sales_2"`);
    expect(text).toContain(`table:display-filter-buttons="true"/></table:database-ranges></office:spreadsheet>`);
  });

  it("bakes the look into generated cell styles and writes structured references as ranges", async () => {
    const model = withTable();
    const before = JSON.stringify(model.styles);
    const xml = await contentOf(await writeOds(model));
    expect(xml).toContain(`table:formula="of:=[.$B4]"`);
    expect(xml).toContain(`table:formula="of:=SUBTOTAL(109;[.$B$2:.$B$4])"`);
    expect(xml).toContain(`table:formula="of:=SUM([.$A$2:.$C$4])"`);
    expect(xml).toContain(`<table:database-range table:name="Sales" table:target-range-address="${model.sheets[0]!.name}.A1:${model.sheets[0]!.name}.C4" table:contains-header="true" table:display-filter-buttons="true"/>`);
    // The header row's cells carry a generated style with the header fill; outside the table, none.
    const header = /<table:table-row[^>]*><table:table-cell table:style-name="(de-ce\d+)"/.exec(xml)?.[1];
    expect(header).toBeDefined();
    expect(new RegExp(`<style:style style:name="${header}"[^>]*>.*?fo:background-color=`).test(xml)).toBe(true);
    expect(JSON.stringify(model.styles)).toBe(before);
    const reread = await open(await writeOds(model));
    expect(reread.sheets[0]!.tables).toBeUndefined();
    expect(reread.sheets[0]!.rows[3]!.cells.find((each) => each.col === 2)?.formula).toBe("$B4");
  });
});
