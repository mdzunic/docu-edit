// The golden wire form of a table created here (SPEC-133 "Tables on the
// wire"): hand-authored against ECMA-376 Part 1 §18.5 (Tables) and written
// BEFORE the serialiser in xlsx/tableXml.ts, which tableXml.spec.ts holds to
// these strings byte for byte. THIS FILE is the normative source for every
// detail the spec left to the golden:
// - the part is the XML declaration and one `table` (§18.5.1.2) binding only
//   the SpreadsheetML namespace; its attributes in schema order: `id` (the
//   workbook-unique number), `name` and `displayName` (both the table name),
//   `ref` (header + data + totals), `headerRowCount="0"` only without a
//   header row, `totalsRowCount="1"` only with a totals row — nothing else
//   (no `totalsRowShown`, whose default says the totals row has been shown);
// - `autoFilter` (§18.3.1.2) with `ref` = header + data rows, only with a
//   header row and filter buttons; no criteria;
// - `tableColumns count` (§18.5.1.4) with one `tableColumn` (§18.5.1.3) per
//   column: `id`, `name`, then `totalsRowFunction` (not for "none") and
//   `totalsRowLabel`; children `calculatedColumnFormula` (§18.5.1.1) and, for
//   a "custom" function, `totalsRowFormula` (§18.5.1.6) — each the formula
//   in the file's spelling (toXlsxFormula of structuredToXlsx: `[#This Row]`),
//   element text escaped; a childless column closes itself;
// - `tableStyleInfo` (§18.5.1.5): `name` (omitted for no style), then the
//   four flags as "1"/"0" in the order Excel writes them.
import type { SheetTable } from "../model";

/** The input of GOLDEN_TABLE_XML: "Sales" at A1:D5 — header row 1, data
 * rows 2–4, totals row 5 — with a label, a SUBTOTAL function, a custom total
 * and a calculated column; TableStyleMedium2, banded rows, filter buttons. */
export const GOLDEN_TABLE: SheetTable = {
  id: "t1",
  name: "Sales",
  ref: { top: 0, left: 0, bottom: 4, right: 3 },
  headerRow: true,
  totalsRow: true,
  columns: [
    { name: "Region", xlsxId: 1, totalsLabel: "Total" },
    { name: "Units", xlsxId: 2, totalsFunction: "sum" },
    { name: "Price", xlsxId: 3, totalsFunction: "custom", totalsFormula: "SUBTOTAL(101,Sales[Price])" },
    { name: "Amount", xlsxId: 4, totalsFunction: "count", calculatedFormula: "Sales[@Units]*Sales[@Price]" },
  ],
  style: { name: "TableStyleMedium2", showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false },
  autoFilter: {},
};

/** GOLDEN_TABLE written as a new part with `id` 1. */
export const GOLDEN_TABLE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:D5" totalsRowCount="1">` +
  `<autoFilter ref="A1:D4"/>` +
  `<tableColumns count="4">` +
  `<tableColumn id="1" name="Region" totalsRowLabel="Total"/>` +
  `<tableColumn id="2" name="Units" totalsRowFunction="sum"/>` +
  `<tableColumn id="3" name="Price" totalsRowFunction="custom"><totalsRowFormula>SUBTOTAL(101,Sales[Price])</totalsRowFormula></tableColumn>` +
  `<tableColumn id="4" name="Amount" totalsRowFunction="count"><calculatedColumnFormula>Sales[[#This Row],[Units]]*Sales[[#This Row],[Price]]</calculatedColumnFormula></tableColumn>` +
  `</tableColumns>` +
  `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>` +
  `</table>`;

/** The input of GOLDEN_HEADERLESS_TABLE_XML: "Rates" at F2:G4 with no header
 * row (so no filter), no totals, no style. */
export const GOLDEN_HEADERLESS_TABLE: SheetTable = {
  id: "t2",
  name: "Rates",
  ref: { top: 1, left: 5, bottom: 3, right: 6 },
  headerRow: false,
  totalsRow: false,
  columns: [
    { name: "Column1", xlsxId: 1 },
    { name: "Column2", xlsxId: 2 },
  ],
  style: { showFirstColumn: false, showLastColumn: false, showRowStripes: false, showColumnStripes: false },
  autoFilter: {},
};

/** GOLDEN_HEADERLESS_TABLE written as a new part with `id` 2. */
export const GOLDEN_HEADERLESS_TABLE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Rates" displayName="Rates" ref="F2:G4" headerRowCount="0">` +
  `<tableColumns count="2"><tableColumn id="1" name="Column1"/><tableColumn id="2" name="Column2"/></tableColumns>` +
  `<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>` +
  `</table>`;

/** ECMA-376 Part 1 §18.3.1.95 `tableParts` (§18.3.1.94 `tablePart`): what a
 * worksheet gaining its first table carries, before `extLst`. */
export const GOLDEN_WORKSHEET_TABLE_PARTS = `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`;

/** The worksheet's relationship to its new table part. */
export const GOLDEN_SHEET_TABLE_RELATIONSHIP = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>`;

/** The `[Content_Types].xml` Override a new table part takes. */
export const GOLDEN_TABLE_OVERRIDE = `<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`;
