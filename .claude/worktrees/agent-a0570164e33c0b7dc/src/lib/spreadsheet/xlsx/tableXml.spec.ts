// SPEC-133's xlsx tables (xlsx/tableXml.ts) per "Tables on the wire" and
// "Canonical and wire forms": reading a table part into a SheetTable (every
// attribute and child the reader models, the file's `[#This Row]` read as
// `@`, what stays out of the model), the worksheet's `<tableParts>` →
// relationship → part walk, golden equality of the new-part serialiser
// against xlsx/tableGolden.ts (and the golden read back), verbatim-preserving
// regeneration of a retained part, the worksheet's `<tableParts>` splice, and
// planTableWrites' package bookkeeping: a clean part verbatim, a dirty or
// disagreeing one regenerated, a new part with its id, relationship and
// `<tableParts>` entry, a deleted table's relationship gone, an unmodellable
// part kept, and a duplicated sheet's closure-copied part claimed.
import { describe, expect, it } from "vitest";
import { relsPathFor } from "../../core/xml";
import type { GridRange } from "../address";
import type { SheetTable, SheetTableStyle } from "../model";
import * as G from "./tableGolden";
import {
  newTablePartXml,
  planTableWrites,
  readTablePart,
  readXlsxSheetTables,
  regenerateTablePart,
  spliceWorksheetTableParts,
  TABLE_CONTENT_TYPE,
  worksheetTablePartIds,
  type TableSheetPlan,
  type TableWritePlan,
} from "./tableXml";

const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const encoder = new TextEncoder();
const decoder = new TextDecoder();

// --- small builders -----------------------------------------------------------------------

const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

/** A table part binding the main namespace: the root's other attributes, then its children. */
const tablePart = (attrs: string, children: string): string => `${DECL}<table xmlns="${MAIN}" ${attrs}>${children}</table>`;
const tableColumns = (...columns: string[]): string => `<tableColumns count="${columns.length}">${columns.join("")}</tableColumns>`;
const MEDIUM2_INFO = `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>`;
const MEDIUM2: SheetTableStyle = { name: "TableStyleMedium2", showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false };

/** The model's reading of a part: a table without `id` and `partName`. */
function modelled(table: SheetTable): Omit<SheetTable, "id"> {
  const { id: _id, partName: _partName, ...rest } = table;
  return rest;
}

/** A worksheet: `<sheetData/>`, `<pageMargins>`, then `tail`. */
const MARGINS = `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`;
const worksheet = (rootAttrs: string, tail: string): string => `${DECL}<worksheet xmlns="${MAIN}"${rootAttrs}><sheetData/>${MARGINS}${tail}</worksheet>`;
const tablePartsXml = (rIds: readonly string[]): string =>
  rIds.length === 0 ? "" : `<tableParts count="${rIds.length}">${rIds.map((rId) => `<tablePart r:id="${rId}"/>`).join("")}</tableParts>`;
/** A worksheet declaring `r` whose `<tableParts>` lists `rIds`. */
const sheetXml = (rIds: readonly string[]): string => worksheet(` xmlns:r="${R}"`, tablePartsXml(rIds));
const rel = (id: string, kind: string, target: string): string => `<Relationship Id="${id}" Type="${R}/${kind}" Target="${target}"/>`;
const rels = (body: string): string => `${DECL}<Relationships xmlns="${PKG}">${body}</Relationships>`;

/** A package of text parts. */
const pkg = (parts: Record<string, string>): Map<string, Uint8Array> => new Map(Object.entries(parts).map(([name, text]) => [name, encoder.encode(text)] as const));
const textOf = (parts: Map<string, Uint8Array>, name: string): string | undefined => {
  const bytes = parts.get(name);
  return bytes === undefined ? undefined : decoder.decode(bytes);
};

/** `ids` minted in order: m1, m2, … */
function minter(): () => string {
  let n = 0;
  return () => {
    n += 1;
    return `m${n}`;
  };
}

const SHEET1 = "xl/worksheets/sheet1.xml";
const SHEET2 = "xl/worksheets/sheet2.xml";
const RELS1 = "xl/worksheets/_rels/sheet1.xml.rels";
const RELS2 = "xl/worksheets/_rels/sheet2.xml.rels";
const TABLE1 = "xl/tables/table1.xml";
const TABLE2 = "xl/tables/table2.xml";
const TABLE3 = "xl/tables/table3.xml";

/** The save's table plan over `parts`, each sheet's `.rels` as the package has it. */
function planOver(parts: Map<string, Uint8Array>, plans: readonly TableSheetPlan[], taken: Set<string> = new Set(parts.keys())): TableWritePlan {
  return planTableWrites(parts, plans, (sheetPart) => textOf(parts, relsPathFor(sheetPart)), taken);
}
const gridPlan = (id: string, partName: string, tables: SheetTable[] | undefined, dirty?: true): TableSheetPlan => ({
  sheet: { id, kind: "grid", ...(tables === undefined ? {} : { tables }), ...(dirty ? { dirty } : {}) },
  partName,
  retained: true,
});
const NOTHING: TableWritePlan = { texts: new Map(), rels: new Map(), edges: new Map(), tableParts: new Map(), added: [] };

// --- the plain "Sales" table: A1:C4, header row, filter buttons ------------------------------

const SALES_XML = tablePart(
  `id="1" name="Sales" displayName="Sales" ref="A1:C4"`,
  `<autoFilter ref="A1:C4"/>` + tableColumns(`<tableColumn id="1" name="Region"/>`, `<tableColumn id="2" name="Price"/>`, `<tableColumn id="3" name="Qty"/>`) + MEDIUM2_INFO,
);
const SALES: SheetTable = {
  id: "m1",
  name: "Sales",
  ref: range(0, 0, 3, 2),
  headerRow: true,
  totalsRow: false,
  columns: [
    { name: "Region", xlsxId: 1 },
    { name: "Price", xlsxId: 2 },
    { name: "Qty", xlsxId: 3 },
  ],
  style: MEDIUM2,
  autoFilter: { sourceXml: `<autoFilter ref="A1:C4"/>` },
  xlsxId: 1,
  partName: TABLE1,
};

/** "Costs" at E1:F3 as a part with `id`. */
const costsXml = (id: number): string =>
  tablePart(`id="${id}" name="Costs" displayName="Costs" ref="E1:F3"`, tableColumns(`<tableColumn id="1" name="Item"/>`, `<tableColumn id="2" name="Cost"/>`) + MEDIUM2_INFO);

/** A new table: "Costs" at A6:B8 with filter buttons. */
const NEW_COSTS: SheetTable = {
  id: "t9",
  name: "Costs",
  ref: range(5, 0, 7, 1),
  headerRow: true,
  totalsRow: false,
  columns: [
    { name: "Item", xlsxId: 1 },
    { name: "Cost", xlsxId: 2 },
  ],
  style: { name: "TableStyleLight1", showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
  autoFilter: {},
};
/** NEW_COSTS as a new part with `id`. */
const newCostsXml = (id: number): string =>
  `${DECL}<table xmlns="${MAIN}" id="${id}" name="Costs" displayName="Costs" ref="A6:B8"><autoFilter ref="A6:B8"/>` +
  `<tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Cost"/></tableColumns>` +
  `<tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;

// --- an Excel-written part: everything the model does not hold ------------------------------

const uid = (n: number): string => `{6E3A1F2B-0C4D-4E5F-8A9B-00000000000${n}}`;
const EXCEL_NS = `xmlns="${MAIN}" xmlns:mc="${MC}" mc:Ignorable="xr xr3" xmlns:xr="${XR}" xmlns:xr3="${XR3}"`;
const SOURCE_FILTER =
  `<autoFilter ref="A1:D4" xr:uid="${uid(8)}">` +
  `<filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn>` +
  `<filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn>` +
  `<filterColumn colId="3"><top10 val="5"/></filterColumn>` +
  `</autoFilter>`;
const SORT_STATE = `<sortState ref="A2:D4" xr:uid="${uid(9)}"><sortCondition descending="1" ref="C2:C4"/></sortState>`;
const COLUMN_EXT = `<extLst><ext uri="{9D7E7F5A-1B2C-4D3E-8F40-000000000001}"/></extLst>`;
const TABLE_EXT = `<extLst><ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}" xmlns:x14="${X14}"><x14:table altTextSummary="Regional sales"/></ext></extLst>`;
const EXCEL_ROOT = `<table ${EXCEL_NS} id="1" xr:uid="${uid(0)}" name="Sales" displayName="Sales" ref="A1:D4" totalsRowShown="0" headerRowDxfId="5" dataDxfId="4">`;
const EXCEL_COLUMNS =
  `<tableColumns count="4">` +
  `<tableColumn id="1" xr3:uid="${uid(1)}" name="Region" dataDxfId="3"/>` +
  `<tableColumn id="2" xr3:uid="${uid(2)}" name="Price" dataDxfId="2"/>` +
  `<tableColumn id="3" xr3:uid="${uid(3)}" name="Qty">${COLUMN_EXT}</tableColumn>` +
  `<tableColumn id="4" xr3:uid="${uid(4)}" name="Amount" dataDxfId="1"><calculatedColumnFormula array="1">Sales[[#This Row],[Price]]*Sales[[#This Row],[Qty]]</calculatedColumnFormula></tableColumn>` +
  `</tableColumns>`;
const EXCEL_PART = `${DECL}${EXCEL_ROOT}${SOURCE_FILTER}${SORT_STATE}${EXCEL_COLUMNS}${MEDIUM2_INFO}${TABLE_EXT}</table>`;
/** EXCEL_PART as the reader models it. */
const EXCEL_TABLE: SheetTable = {
  id: "t1",
  name: "Sales",
  ref: range(0, 0, 3, 3),
  headerRow: true,
  totalsRow: false,
  columns: [
    { name: "Region", xlsxId: 1 },
    { name: "Price", xlsxId: 2 },
    { name: "Qty", xlsxId: 3 },
    { name: "Amount", xlsxId: 4, calculatedFormula: "Sales[@Price]*Sales[@Qty]" },
  ],
  style: MEDIUM2,
  autoFilter: { sourceXml: SOURCE_FILTER },
  xlsxId: 1,
  partName: TABLE1,
};

// --- reading --------------------------------------------------------------------------------

describe("readTablePart", () => {
  it("reads a plain part into the model's table", () => {
    expect(readTablePart(SALES_XML)).toEqual(modelled(SALES));
  });

  it("takes displayName as the name, falling back to name; reads ref and table@id → xlsxId", () => {
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`);
    const both = readTablePart(tablePart(`id="7" name="Table1" displayName="Sales" ref="B2:D6"`, columns))!;
    expect([both.name, both.ref, both.xlsxId]).toEqual(["Sales", range(1, 1, 5, 3), 7]);
    expect(readTablePart(tablePart(`id="7" name="Table1" ref="B2:D6"`, columns))!.name).toBe("Table1");
  });

  it("reads headerRowCount 0 as no header row and an absent one as a header row", () => {
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" headerRowCount="0"`, columns))!.headerRow).toBe(false);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3"`, columns))!.headerRow).toBe(true);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" headerRowCount="1"`, columns))!.headerRow).toBe(true);
  });

  it("reads totalsRowCount 1 as a totals row; absent or 0 as none (totalsRowShown alone is not one)", () => {
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" totalsRowCount="1"`, columns))!.totalsRow).toBe(true);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3"`, columns))!.totalsRow).toBe(false);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" totalsRowCount="0" totalsRowShown="1"`, columns))!.totalsRow).toBe(false);
  });

  it("reads tableStyleInfo's name and show* flags as \"1\" or \"true\"", () => {
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`);
    const styled = readTablePart(
      tablePart(
        `id="1" name="T" displayName="T" ref="A1:A3"`,
        columns + `<tableStyleInfo name="TableStyleDark3" showFirstColumn="true" showLastColumn="1" showRowStripes="0" showColumnStripes="false"/>`,
      ),
    )!;
    expect(styled.style).toEqual({ name: "TableStyleDark3", showFirstColumn: true, showLastColumn: true, showRowStripes: false, showColumnStripes: false });
    const plain = readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3"`, columns + `<tableStyleInfo showRowStripes="1"/>`))!;
    expect(plain.style).toEqual({ showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false });
    const none = readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3"`, columns))!;
    expect(none.style).toEqual({ showFirstColumn: false, showLastColumn: false, showRowStripes: false, showColumnStripes: false });
  });

  it("reads each tableColumn: name (_xHHHH_ decoded), id → xlsxId, totalsRowFunction (none → absent), totalsRowLabel", () => {
    const read = readTablePart(
      tablePart(
        `id="1" name="T" displayName="T" ref="A1:D4" totalsRowCount="1"`,
        tableColumns(
          `<tableColumn id="1" name="Region" totalsRowLabel="Grand total"/>`,
          `<tableColumn id="2" name="Unit_x0020_Price" totalsRowFunction="none"/>`,
          `<tableColumn id="5" name="Qty" totalsRowFunction="countNums"/>`,
          `<tableColumn id="6" name="Amount" totalsRowFunction="stdDev"/>`,
        ),
      ),
    )!;
    expect(read.columns).toEqual([
      { name: "Region", xlsxId: 1, totalsLabel: "Grand total" },
      { name: "Unit Price", xlsxId: 2 },
      { name: "Qty", xlsxId: 5, totalsFunction: "countNums" },
      { name: "Amount", xlsxId: 6, totalsFunction: "stdDev" },
    ]);
  });

  it("reads calculatedColumnFormula and totalsRowFormula through fromXlsxFormula + structuredFromXlsx", () => {
    const read = readTablePart(
      tablePart(
        `id="1" name="Sales" displayName="Sales" ref="A1:C5" totalsRowCount="1"`,
        tableColumns(
          `<tableColumn id="1" name="Price"/>`,
          `<tableColumn id="2" name="Qty" totalsRowFunction="custom"><totalsRowFormula>SUBTOTAL(109,Sales[Qty])*2</totalsRowFormula></tableColumn>`,
          `<tableColumn id="3" name="Amount"><calculatedColumnFormula>Sales[[#This Row],[Price]]*Sales[[#This Row],[Qty]]</calculatedColumnFormula></tableColumn>`,
        ),
      ),
    )!;
    expect(read.columns).toEqual([
      { name: "Price", xlsxId: 1 },
      { name: "Qty", xlsxId: 2, totalsFunction: "custom", totalsFormula: "SUBTOTAL(109,Sales[Qty])*2" },
      { name: "Amount", xlsxId: 3, calculatedFormula: "Sales[@Price]*Sales[@Qty]" },
    ]);
  });

  it("leaves a column formula that does not parse out of the model", () => {
    const read = readTablePart(
      tablePart(`id="1" name="T" displayName="T" ref="A1:B3"`, tableColumns(`<tableColumn id="1" name="A"/>`, `<tableColumn id="2" name="B"><calculatedColumnFormula>SUM(</calculatedColumnFormula></tableColumn>`)),
    )!;
    expect(read.columns).toEqual([
      { name: "A", xlsxId: 1 },
      { name: "B", xlsxId: 2 },
    ]);
  });

  it("keeps the table's autoFilter verbatim as autoFilter.sourceXml; no autoFilter → no filter buttons", () => {
    const filter = `<autoFilter ref="A1:B3"><filterColumn colId="1"><filters><filter val="East"/></filters></filterColumn></autoFilter>`;
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`, `<tableColumn id="2" name="B"/>`);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:B3"`, filter + columns))!.autoFilter).toEqual({ sourceXml: filter });
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:B3"`, columns))!.autoFilter).toBeUndefined();
  });

  it("locks a queryTable or xml table; a worksheet table is not locked", () => {
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" tableType="queryTable"`, columns))!.locked).toBe(true);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" tableType="xml"`, columns))!.locked).toBe(true);
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="A1:A3" tableType="worksheet"`, columns))!.locked).toBeUndefined();
  });

  it("reads Excel's part with everything unmodelled into the model's table", () => {
    expect(readTablePart(EXCEL_PART)).toEqual(modelled(EXCEL_TABLE));
  });

  it("is null for a part whose root or ref does not read", () => {
    const columns = tableColumns(`<tableColumn id="1" name="A"/>`);
    expect(readTablePart(`${DECL}<table xmlns="${MAIN}" id="1" name="T" ref="A1:A3">`)).toBeNull();
    expect(readTablePart(`${DECL}<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`)).toBeNull();
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T" ref="nonsense"`, columns))).toBeNull();
    expect(readTablePart(tablePart(`id="1" name="T" displayName="T"`, columns))).toBeNull();
  });
});

describe("worksheetTablePartIds", () => {
  it("lists <tableParts>' relationship ids in order, under whatever prefix the root binds", () => {
    expect(worksheetTablePartIds(sheetXml(["rId3", "rId1"]))).toEqual(["rId3", "rId1"]);
    expect(worksheetTablePartIds(worksheet(` xmlns:rel="${R}"`, `<tableParts count="1"><tablePart rel:id="rId2"/></tableParts>`))).toEqual(["rId2"]);
  });

  it("is undefined for a worksheet without <tableParts>", () => {
    expect(worksheetTablePartIds(sheetXml([]))).toBeUndefined();
  });
});

describe("readXlsxSheetTables", () => {
  it("follows <tableParts> → the sheet's table relationship → the part, setting partName and minting ids", () => {
    const parts = pkg({
      [SHEET1]: sheetXml(["rId2", "rId1"]),
      [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "table", "../tables/table2.xml")),
      [TABLE1]: SALES_XML,
      [TABLE2]: costsXml(2),
    });
    const tables = readXlsxSheetTables(parts, SHEET1, textOf(parts, SHEET1)!, minter());
    expect(tables.map((table) => [table.id, table.name, table.partName, table.xlsxId])).toEqual([
      ["m1", "Costs", TABLE2, 2],
      ["m2", "Sales", TABLE1, 1],
    ]);
    expect(tables[1]).toEqual({ ...SALES, id: "m2" });
  });

  it("skips a part it cannot model, a relationship of another kind and a missing one", () => {
    const parts = pkg({
      [SHEET1]: sheetXml(["rId1", "rId2", "rId3", "rId9"]),
      [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "table", "../tables/table2.xml") + rel("rId3", "drawing", "../drawings/drawing1.xml")),
      [TABLE1]: tablePart(`id="1" name="Broken" displayName="Broken" ref="not a range"`, tableColumns(`<tableColumn id="1" name="A"/>`)),
      [TABLE2]: SALES_XML,
      "xl/drawings/drawing1.xml": `${DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`,
    });
    const tables = readXlsxSheetTables(parts, SHEET1, textOf(parts, SHEET1)!, minter());
    expect(tables).toEqual([{ ...SALES, id: "m1", partName: TABLE2 }]);
  });

  it("reads no tables from a worksheet without <tableParts>", () => {
    const parts = pkg({ [SHEET1]: sheetXml([]), [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml")), [TABLE1]: SALES_XML });
    expect(readXlsxSheetTables(parts, SHEET1, textOf(parts, SHEET1)!, minter())).toEqual([]);
  });
});

// --- the golden --------------------------------------------------------------------------------

describe("the new part's golden (tableGolden.ts)", () => {
  it("writes GOLDEN_TABLE with id 1 byte for byte", () => {
    expect(newTablePartXml(G.GOLDEN_TABLE, 1)).toBe(G.GOLDEN_TABLE_XML);
  });

  it("writes GOLDEN_HEADERLESS_TABLE with id 2 byte for byte", () => {
    expect(newTablePartXml(G.GOLDEN_HEADERLESS_TABLE, 2)).toBe(G.GOLDEN_HEADERLESS_TABLE_XML);
  });

  it("reads each golden back as its model table", () => {
    const { id: _id, ...table } = G.GOLDEN_TABLE;
    expect(readTablePart(G.GOLDEN_TABLE_XML)).toEqual({ ...table, xlsxId: 1, autoFilter: { sourceXml: `<autoFilter ref="A1:D4"/>` } });
    // No header row, so no filter buttons on the wire.
    const { id: _headerlessId, autoFilter: _filter, ...headerless } = G.GOLDEN_HEADERLESS_TABLE;
    expect(readTablePart(G.GOLDEN_HEADERLESS_TABLE_XML)).toEqual({ ...headerless, xlsxId: 2 });
  });

  it("writes the read-back golden as the golden again", () => {
    expect(newTablePartXml({ id: "x", ...readTablePart(G.GOLDEN_TABLE_XML)! }, 1)).toBe(G.GOLDEN_TABLE_XML);
    expect(newTablePartXml({ id: "x", ...readTablePart(G.GOLDEN_HEADERLESS_TABLE_XML)! }, 2)).toBe(G.GOLDEN_HEADERLESS_TABLE_XML);
  });

  it("has the table part's content type", () => {
    expect(TABLE_CONTENT_TYPE).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
  });
});

// --- regenerating a retained part ----------------------------------------------------------

/** An autoFilter at `ref` with criteria on its second column. */
const filterOnB = (ref: string): string => `<autoFilter ref="${ref}"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter>`;

describe("regenerateTablePart", () => {
  it("gives back the part byte for byte while the model still says what it says", () => {
    expect(regenerateTablePart(EXCEL_PART, { ...EXCEL_TABLE, dirty: true }, 1)).toBe(EXCEL_PART);
    expect(regenerateTablePart(SALES_XML, { ...SALES, dirty: true }, 1)).toBe(SALES_XML);
  });

  it("keeps every unmodelled attribute and child through a rename, a new ref, columns removed and added, totals and style options", () => {
    const edited: SheetTable = {
      ...EXCEL_TABLE,
      name: "Revenue",
      ref: range(0, 1, 5, 5),
      totalsRow: true,
      columns: [
        { name: "Price", xlsxId: 2, totalsLabel: "Total" },
        { name: "Qty", xlsxId: 3, totalsFunction: "sum" },
        { name: "Amount", xlsxId: 4, calculatedFormula: "Revenue[@Price]*Revenue[@Qty]" },
        { name: "Discount" },
        { name: "Net", calculatedFormula: "Revenue[@Amount]-Revenue[@Discount]" },
      ],
      style: { name: "TableStyleLight9", showFirstColumn: true, showLastColumn: false, showRowStripes: false, showColumnStripes: true },
      dirty: true,
    };
    expect(regenerateTablePart(EXCEL_PART, edited, 6)).toBe(
      DECL +
        `<table ${EXCEL_NS} id="6" xr:uid="${uid(0)}" name="Revenue" displayName="Revenue" ref="B1:F6" headerRowDxfId="5" dataDxfId="4" totalsRowCount="1">` +
        // ref = header + data; Region's filter goes, Qty's and Amount's follow their columns.
        `<autoFilter ref="B1:F5" xr:uid="${uid(8)}">` +
        `<filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn>` +
        `<filterColumn colId="2"><top10 val="5"/></filterColumn>` +
        `</autoFilter>` +
        SORT_STATE +
        `<tableColumns count="5">` +
        `<tableColumn id="2" xr3:uid="${uid(2)}" name="Price" dataDxfId="2" totalsRowLabel="Total"/>` +
        `<tableColumn id="3" xr3:uid="${uid(3)}" name="Qty" totalsRowFunction="sum">${COLUMN_EXT}</tableColumn>` +
        `<tableColumn id="4" xr3:uid="${uid(4)}" name="Amount" dataDxfId="1"><calculatedColumnFormula array="1">Revenue[[#This Row],[Price]]*Revenue[[#This Row],[Qty]]</calculatedColumnFormula></tableColumn>` +
        `<tableColumn id="5" name="Discount"/>` +
        `<tableColumn id="6" name="Net"><calculatedColumnFormula>Revenue[[#This Row],[Amount]]-Revenue[[#This Row],[Discount]]</calculatedColumnFormula></tableColumn>` +
        `</tableColumns>` +
        `<tableStyleInfo name="TableStyleLight9" showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="1"/>` +
        TABLE_EXT +
        `</table>`,
    );
  });

  it("omits the autoFilter without filter buttons, keeping sortState and extLst", () => {
    expect(regenerateTablePart(EXCEL_PART, { ...EXCEL_TABLE, autoFilter: undefined, dirty: true }, 1)).toBe(EXCEL_PART.replace(SOURCE_FILTER, ""));
  });

  it("omits the autoFilter without a header row, writing headerRowCount 0", () => {
    const headerless: SheetTable = { ...EXCEL_TABLE, headerRow: false, ref: range(1, 0, 3, 3), dirty: true };
    expect(regenerateTablePart(EXCEL_PART, headerless, 1)).toBe(
      EXCEL_PART.replace(`ref="A1:D4" totalsRowShown="0" headerRowDxfId="5" dataDxfId="4">`, `ref="A2:D4" totalsRowShown="0" headerRowDxfId="5" dataDxfId="4" headerRowCount="0">`).replace(SOURCE_FILTER, ""),
    );
  });

  it("adds <autoFilter ref/> before the columns when the source had none, dropping headerRowCount 0 for a header row", () => {
    const source = tablePart(
      `id="2" name="Rates" displayName="Rates" ref="F2:G4" headerRowCount="0" dataDxfId="7"`,
      tableColumns(`<tableColumn id="1" name="Column1"/>`, `<tableColumn id="2" name="Column2"/>`) + MEDIUM2_INFO,
    );
    const table: SheetTable = {
      id: "t2",
      name: "Rates",
      ref: range(0, 5, 3, 6),
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "Column1", xlsxId: 1 },
        { name: "Column2", xlsxId: 2 },
      ],
      style: MEDIUM2,
      autoFilter: {},
      dirty: true,
    };
    expect(regenerateTablePart(source, table, 2)).toBe(
      tablePart(
        `id="2" name="Rates" displayName="Rates" ref="F1:G4" dataDxfId="7"`,
        `<autoFilter ref="F1:G4"/>` + tableColumns(`<tableColumn id="1" name="Column1"/>`, `<tableColumn id="2" name="Column2"/>`) + MEDIUM2_INFO,
      ),
    );
  });

  it("takes the model's autoFilter source when the part has none, its ref patched", () => {
    const source = tablePart(`id="4" name="T" displayName="T" ref="A1:B3"`, tableColumns(`<tableColumn id="1" name="A"/>`, `<tableColumn id="2" name="B"/>`) + MEDIUM2_INFO);
    const table: SheetTable = { id: "t4", ...readTablePart(source)!, ref: range(0, 0, 4, 1), autoFilter: { sourceXml: filterOnB("A1:B3") }, dirty: true };
    expect(regenerateTablePart(source, table, 4)).toBe(
      tablePart(`id="4" name="T" displayName="T" ref="A1:B5"`, filterOnB("A1:B5") + tableColumns(`<tableColumn id="1" name="A"/>`, `<tableColumn id="2" name="B"/>`) + MEDIUM2_INFO),
    );
  });

  it("toggles the totals row: totalsRowCount and totalsRowShown on the root, the filter over header + data only", () => {
    /** "T" with a totals label and a custom total; `rootAttrs` after `displayName`. */
    const part = (rootAttrs: string, filterRef: string): string =>
      tablePart(
        `id="3" name="T" displayName="T" ${rootAttrs}`,
        `<autoFilter ref="${filterRef}"/>` +
          tableColumns(
            `<tableColumn id="1" name="Item" totalsRowLabel="Total"/>`,
            `<tableColumn id="2" name="Cost" totalsRowFunction="custom" dataDxfId="1"><totalsRowFormula>SUBTOTAL(109,T[Cost])</totalsRowFormula></tableColumn>`,
          ) +
          MEDIUM2_INFO,
      );
    const shown = part(`ref="A1:B5" totalsRowCount="1" headerRowDxfId="2"`, "A1:B4");
    const table: SheetTable = { id: "t3", ...readTablePart(shown)!, dirty: true };
    // Off: the root loses totalsRowCount; the columns keep their totals settings.
    expect(regenerateTablePart(shown, { ...table, totalsRow: false, ref: range(0, 0, 3, 1) }, 3)).toBe(part(`ref="A1:B4" headerRowDxfId="2"`, "A1:B4"));
    // On, from a part whose totals row was never shown: totalsRowShown="0" goes.
    const never = part(`ref="A1:B4" totalsRowShown="0" headerRowDxfId="2"`, "A1:B4");
    expect(regenerateTablePart(never, { ...table, totalsRow: true, ref: range(0, 0, 4, 1) }, 3)).toBe(part(`ref="A1:B5" headerRowDxfId="2" totalsRowCount="1"`, "A1:B4"));
  });

  it("keeps a column formula the reader could not model, and a formula child the model still agrees with", () => {
    const source = tablePart(
      `id="1" name="T" displayName="T" ref="A1:B3"`,
      tableColumns(
        `<tableColumn id="1" name="A"><calculatedColumnFormula array="1">T[[#This Row],[B]]*2</calculatedColumnFormula></tableColumn>`,
        `<tableColumn id="2" name="B"><calculatedColumnFormula>SUM(</calculatedColumnFormula></tableColumn>`,
      ) + MEDIUM2_INFO,
    );
    const read = readTablePart(source)!;
    expect(read.columns).toEqual([
      { name: "A", xlsxId: 1, calculatedFormula: "T[@B]*2" },
      { name: "B", xlsxId: 2 },
    ]);
    expect(regenerateTablePart(source, { id: "t", ...read, ref: range(0, 0, 4, 1), dirty: true }, 1)).toBe(source.replace(`ref="A1:B3"`, `ref="A1:B5"`));
  });

  it("writes new elements under the source's prefix", () => {
    const source = `${DECL}<x:table xmlns:x="${MAIN}" id="1" name="T" displayName="T" ref="A1:A3"><x:tableColumns count="1"><x:tableColumn id="1" name="A"/></x:tableColumns></x:table>`;
    const table: SheetTable = {
      id: "t",
      name: "T",
      ref: range(0, 0, 2, 1),
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "A", xlsxId: 1 }, { name: "B" }],
      style: MEDIUM2,
      autoFilter: {},
      dirty: true,
    };
    expect(regenerateTablePart(source, table, 1)).toBe(
      `${DECL}<x:table xmlns:x="${MAIN}" id="1" name="T" displayName="T" ref="A1:B3"><x:autoFilter ref="A1:B3"/>` +
        `<x:tableColumns count="2"><x:tableColumn id="1" name="A"/><x:tableColumn id="2" name="B"/></x:tableColumns>` +
        `<x:tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></x:table>`,
    );
  });
});

// --- the worksheet's <tableParts> --------------------------------------------------------------

describe("spliceWorksheetTableParts", () => {
  const SHEET_EXT = `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}"/></extLst>`;
  const R_ATTR = ` xmlns:r="${R}"`;

  it("inserts <tableParts> before extLst (CT_Worksheet order)", () => {
    expect(spliceWorksheetTableParts(worksheet(R_ATTR, `<drawing r:id="rId1"/>${SHEET_EXT}`), ["rId2", "rId3"])).toBe(
      worksheet(R_ATTR, `<drawing r:id="rId1"/><tableParts count="2"><tablePart r:id="rId2"/><tablePart r:id="rId3"/></tableParts>${SHEET_EXT}`),
    );
  });

  it("appends <tableParts> at the root's end without extLst — the golden for a first table", () => {
    expect(spliceWorksheetTableParts(sheetXml([]), ["rId1"])).toBe(worksheet(R_ATTR, G.GOLDEN_WORKSHEET_TABLE_PARTS));
  });

  it("replaces an existing <tableParts> in place", () => {
    expect(spliceWorksheetTableParts(worksheet(R_ATTR, `${tablePartsXml(["rId2", "rId3"])}${SHEET_EXT}`), ["rId3"])).toBe(worksheet(R_ATTR, `${tablePartsXml(["rId3"])}${SHEET_EXT}`));
  });

  it("removes <tableParts> for no ids, and leaves a sheet without one alone", () => {
    expect(spliceWorksheetTableParts(worksheet(R_ATTR, `${tablePartsXml(["rId2"])}${SHEET_EXT}`), [])).toBe(worksheet(R_ATTR, SHEET_EXT));
    expect(spliceWorksheetTableParts(sheetXml([]), [])).toBe(sheetXml([]));
  });

  it("uses the root's relationships prefix, or declares xmlns:r when it has none", () => {
    expect(spliceWorksheetTableParts(worksheet(` xmlns:rel="${R}"`, ""), ["rId1"])).toBe(
      worksheet(` xmlns:rel="${R}"`, `<tableParts count="1"><tablePart rel:id="rId1"/></tableParts>`),
    );
    expect(spliceWorksheetTableParts(worksheet("", ""), ["rId1"])).toBe(worksheet("", `<tableParts count="1" xmlns:r="${R}"><tablePart r:id="rId1"/></tableParts>`));
  });
});

// --- the save's plan ------------------------------------------------------------------------

describe("planTableWrites", () => {
  /** Sheet1 relating Sales (table1.xml) as rId1. */
  const salesBook = (extra: Record<string, string> = {}): Map<string, Uint8Array> =>
    pkg({ [SHEET1]: sheetXml(["rId1"]), [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml")), [TABLE1]: SALES_XML, ...extra });
  const reopened = (parts: Map<string, Uint8Array>): SheetTable[] => readXlsxSheetTables(parts, SHEET1, textOf(parts, SHEET1)!, minter());

  it("writes nothing for a clean table whose part agrees — its part stays verbatim", () => {
    const parts = salesBook();
    expect(planOver(parts, [gridPlan("s1", SHEET1, reopened(parts))])).toEqual(NOTHING);
  });

  it("regenerates a dirty table's part", () => {
    const parts = salesBook();
    const plan = planOver(parts, [gridPlan("s1", SHEET1, [{ ...SALES, name: "Revenue", dirty: true }])]);
    expect(plan).toEqual({ ...NOTHING, texts: new Map([[TABLE1, SALES_XML.replace(`name="Sales" displayName="Sales"`, `name="Revenue" displayName="Revenue"`)]]) });
  });

  it("regenerates a clean table whose ref no longer agrees with its part", () => {
    const parts = salesBook();
    const plan = planOver(parts, [gridPlan("s1", SHEET1, [{ ...SALES, ref: range(0, 0, 5, 2) }])]);
    expect(plan).toEqual({ ...NOTHING, texts: new Map([[TABLE1, SALES_XML.replaceAll(`ref="A1:C4"`, `ref="A1:C6"`)]]) });
  });

  it("writes a new table at the next free part with a workbook-unique id, a relationship and a <tableParts> entry", () => {
    const parts = pkg({
      [SHEET1]: sheetXml(["rId1"]),
      [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "drawing", "../drawings/drawing1.xml")),
      [TABLE1]: SALES_XML,
      [SHEET2]: sheetXml(["rId1"]),
      [RELS2]: rels(rel("rId1", "table", "../tables/table3.xml")),
      [TABLE3]: costsXml(7),
    });
    const taken = new Set(parts.keys());
    const plan = planOver(parts, [gridPlan("s1", SHEET1, [SALES, NEW_COSTS]), gridPlan("s2", SHEET2, undefined)], taken);
    const TABLE4 = "xl/tables/table4.xml";
    expect(plan).toEqual({
      texts: new Map([[TABLE4, newCostsXml(8)]]),
      rels: new Map([[RELS1, rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "drawing", "../drawings/drawing1.xml") + rel("rId3", "table", "../tables/table4.xml"))]]),
      edges: new Map([[SHEET1, [TABLE1, "xl/drawings/drawing1.xml", TABLE4]]]),
      tableParts: new Map([[SHEET1, ["rId1", "rId3"]]]),
      added: [TABLE4],
    });
    expect(taken.has(TABLE4)).toBe(true);
  });

  it("gives a sheet's first table the golden part, relationship, <tableParts> and Override", () => {
    const parts = pkg({ [SHEET1]: sheetXml([]) });
    const plan = planOver(parts, [gridPlan("s1", SHEET1, [G.GOLDEN_TABLE])]);
    expect(plan).toEqual({
      texts: new Map([[TABLE1, G.GOLDEN_TABLE_XML]]),
      rels: new Map([[RELS1, rels(G.GOLDEN_SHEET_TABLE_RELATIONSHIP)]]),
      edges: new Map([[SHEET1, [TABLE1]]]),
      tableParts: new Map([[SHEET1, ["rId1"]]]),
      added: [TABLE1],
    });
    expect(spliceWorksheetTableParts(sheetXml([]), plan.tableParts.get(SHEET1)!)).toBe(sheetXml(["rId1"]));
    expect(sheetXml(["rId1"])).toContain(G.GOLDEN_WORKSHEET_TABLE_PARTS);
    expect(`<Override PartName="/${plan.added[0]}" ContentType="${TABLE_CONTENT_TYPE}"/>`).toBe(G.GOLDEN_TABLE_OVERRIDE);
  });

  it("drops a deleted table's relationship and rewrites <tableParts> without it", () => {
    const parts = pkg({
      [SHEET1]: sheetXml(["rId1", "rId2"]),
      [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "table", "../tables/table2.xml")),
      [TABLE1]: SALES_XML,
      [TABLE2]: costsXml(2),
    });
    expect(planOver(parts, [gridPlan("s1", SHEET1, [SALES])])).toEqual({
      ...NOTHING,
      rels: new Map([[RELS1, rels(rel("rId1", "table", "../tables/table1.xml"))]]),
      edges: new Map([[SHEET1, [TABLE1]]]),
      tableParts: new Map([[SHEET1, ["rId1"]]]),
    });
  });

  it("drops the last table's relationship and <tableParts> for a dirty sheet without tables", () => {
    const parts = salesBook();
    expect(planOver(parts, [gridPlan("s1", SHEET1, undefined, true)])).toEqual({
      ...NOTHING,
      rels: new Map([[RELS1, rels("")]]),
      edges: new Map([[SHEET1, []]]),
      tableParts: new Map([[SHEET1, []]]),
    });
  });

  it("keeps a part the reader cannot model, listed after the model's tables", () => {
    const broken = tablePart(`id="2" name="Broken" displayName="Broken" ref="not a range"`, tableColumns(`<tableColumn id="1" name="A"/>`));
    const parts = pkg({
      [SHEET1]: sheetXml(["rId2", "rId1"]),
      [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "table", "../tables/table2.xml")),
      [TABLE1]: SALES_XML,
      [TABLE2]: broken,
    });
    expect(planOver(parts, [gridPlan("s1", SHEET1, [SALES])])).toEqual({ ...NOTHING, tableParts: new Map([[SHEET1, ["rId1", "rId2"]]]) });
    // A new table joins the model's, before the kept part; its id passes the kept part's.
    const plan = planOver(parts, [gridPlan("s1", SHEET1, [SALES, NEW_COSTS])]);
    expect(plan.texts).toEqual(new Map([[TABLE3, newCostsXml(3)]]));
    expect(plan.added).toEqual([TABLE3]);
    expect(plan.rels.get(RELS1)).toBe(rels(rel("rId1", "table", "../tables/table1.xml") + rel("rId2", "table", "../tables/table2.xml") + rel("rId3", "table", "../tables/table3.xml")));
    expect(plan.tableParts).toEqual(new Map([[SHEET1, ["rId1", "rId3", "rId2"]]]));
  });

  it("moves a table's part to the sheet it moved to: that sheet gains the relationship, the old one loses it", () => {
    const parts = salesBook({ [SHEET2]: sheetXml([]), [RELS2]: rels(rel("rId1", "drawing", "../drawings/drawing1.xml")) });
    expect(planOver(parts, [gridPlan("s1", SHEET1, []), gridPlan("s2", SHEET2, [SALES])])).toEqual({
      ...NOTHING,
      rels: new Map([
        [RELS1, rels("")],
        [RELS2, rels(rel("rId1", "drawing", "../drawings/drawing1.xml") + rel("rId2", "table", "../tables/table1.xml"))],
      ]),
      edges: new Map([
        [SHEET1, []],
        [SHEET2, ["xl/drawings/drawing1.xml", TABLE1]],
      ]),
      tableParts: new Map([
        [SHEET1, []],
        [SHEET2, ["rId2"]],
      ]),
    });
  });

  describe("a duplicated sheet", () => {
    const ORIGIN_XML = tablePart(
      `id="1" name="Sales" displayName="Sales" ref="A1:C4" headerRowDxfId="3"`,
      `<autoFilter ref="A1:C4"/>` +
        tableColumns(`<tableColumn id="1" name="Region" dataDxfId="2"/>`, `<tableColumn id="2" name="Price"/>`, `<tableColumn id="3" name="Qty" dataDxfId="1"/>`) +
        MEDIUM2_INFO,
    );
    /** SPEC-121's closure copy of ORIGIN_XML for the copy: id max + 1, `<name>_<n>`. */
    const COPY_XML = ORIGIN_XML.replace(`id="1" name="Sales" displayName="Sales"`, `id="2" name="Sales_2" displayName="Sales_2"`);
    const book = (): Map<string, Uint8Array> =>
      pkg({
        [SHEET1]: sheetXml(["rId1"]),
        [RELS1]: rels(rel("rId1", "table", "../tables/table1.xml")),
        [TABLE1]: ORIGIN_XML,
        [SHEET2]: sheetXml(["rId1"]),
        [RELS2]: rels(rel("rId1", "table", "../tables/table2.xml")),
        [TABLE2]: COPY_XML,
      });
    const origin = (): SheetTable => ({ id: "t1", ...readTablePart(ORIGIN_XML)!, partName: TABLE1 });

    it("lets the copy's table claim its closure-copied part: regenerated in place, no new part, relationship kept", () => {
      const parts = book();
      const copy: SheetTable = { ...origin(), id: "c1", name: "Sales_2", xlsxId: 2, partName: TABLE2, dirty: true };
      const plan = planOver(parts, [gridPlan("s1", SHEET1, [origin()]), gridPlan("s2", SHEET2, [copy], true)]);
      expect(plan).toEqual({ ...NOTHING, texts: new Map([[TABLE2, COPY_XML]]) });
      // Renamed after the copy: still that part, its unmodelled attributes kept.
      const renamed = planOver(parts, [gridPlan("s1", SHEET1, [origin()]), gridPlan("s2", SHEET2, [{ ...copy, name: "Copy" }], true)]);
      expect(renamed).toEqual({ ...NOTHING, texts: new Map([[TABLE2, COPY_XML.replace(`name="Sales_2" displayName="Sales_2"`, `name="Copy" displayName="Copy"`)]]) });
    });

    it("drops a closure-copied part no copy's table claims, writing the copy's table as a new part", () => {
      const parts = book();
      const copy: SheetTable = { ...origin(), id: "c1", name: "Sales_2", xlsxId: 2, dirty: true };
      delete copy.partName;
      const plan = planOver(parts, [gridPlan("s1", SHEET1, [origin()]), gridPlan("s2", SHEET2, [copy], true)]);
      expect(plan.added).toEqual([TABLE3]);
      // The dropped copy (id 2) still counts while the plan is made: max + 1.
      expect(plan.texts).toEqual(new Map([[TABLE3, newTablePartXml(copy, 3)]]));
      expect(plan.rels).toEqual(new Map([[RELS2, rels(rel("rId2", "table", "../tables/table3.xml"))]]));
      expect(plan.tableParts).toEqual(new Map([[SHEET2, ["rId2"]]]));
    });
  });
});
