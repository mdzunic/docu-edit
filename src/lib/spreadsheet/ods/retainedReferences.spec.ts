// SPEC-124 "Retained references", the .ods half (ods/retainedReferences.ts):
// every row of the ods table moved by a structural edit through
// structuralMapper, asserted on the adjusted text and with the span-diff
// helper (only attribute values change, plus the elements and attributes the
// rules remove); entry dropping and element removal, the formula brackets
// (entities and strings around them kept, `#REF!`, external and 3-D
// references left alone), hosts (a renamed sheet, a duplicated sheet's own
// text, an embedded chart's sheet), tracked changes, content.xml = prefix +
// tables + suffix after every edit, copy-on-write identity and the returned
// sheet ids. SPEC-128: a database range that is a sheet's autofilter is left
// alone (the model's adjuster moves it), every other one still moves.
import { describe, expect, it, vi } from "vitest";
import { base64ToBytes, bytesToBase64, type OpaquePart } from "../../core/packageParts";
import { scanXml } from "../../core/xmlScanner";
import { changedSpans } from "../fixtures/spanDiff";
import type { Row, Sheet, SpreadsheetModel } from "../model";
import { SPREADSHEET_READER_VERSION } from "../open";
import { rewriteRetainedReferences } from "../retainedReferences";
import { structuralMapper, type StructuralEdit } from "../structure";
import { decodeXmlText } from "../xmlSplice";
import { sourceTables } from "./package";
import { rewriteOdsRetainedReferences } from "./retainedReferences";

// scanXml as it is, counted (the prefilter test).
vi.mock("../../core/xmlScanner", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../core/xmlScanner")>();
  return { ...actual, scanXml: vi.fn(actual.scanXml) };
});

// --- a synthetic LibreOffice-shaped workbook ------------------------------------------------

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"',
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
].join(" ");
const HEAD = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS} office:version="1.3"><office:automatic-styles/><office:body><office:spreadsheet>`;
const TAIL = `\n</office:spreadsheet></office:body></office:document-content>`;
const COLUMNS = `<table:table-column table:number-columns-repeated="8"/>`;
const EMPTY_ROWS = `<table:table-row table:number-rows-repeated="20"><table:table-cell table:number-columns-repeated="8"/></table:table-row>`;

interface TableSpec {
  /** `table:name` as written (escaped). */
  name: string;
  /** The model's name; the decoded `table:name` when absent. */
  modelName?: string;
  attrs?: string;
  leading?: string;
  rows?: string;
  trailing?: string;
  cells?: Row[];
  /** A deleted sheet's table: in content.xml, not in the model. */
  deleted?: true;
}

interface CopySpec {
  name: string;
  copiedFrom: string;
  sourceXml: string;
  /** The copy shares its origin's rows (by reference) when absent. */
  cells?: Row[];
}

interface BookSpec {
  prefix?: string;
  suffix?: string;
  tables: TableSpec[];
  copies?: CopySpec[];
  /** Further parts: name → text. */
  parts?: Record<string, string>;
  /** content.xml written with a UTF-8 byte-order mark. */
  bom?: boolean;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function tableXml(spec: Omit<TableSpec, "modelName" | "cells">): string {
  return `\n  <table:table table:name="${spec.name}" table:style-name="ta1"${spec.attrs ?? ""}>${spec.leading ?? ""}${COLUMNS}${spec.rows ?? EMPTY_ROWS}${spec.trailing ?? ""}</table:table>`;
}

function grid(name: string, sourceXml: string, rows: Row[]): Sheet {
  return {
    conditionalFormats: [],
    id: `sheet:${name}`,
    name,
    visibility: "visible",
    kind: "grid",
    rows,
    sourceXml,
    columns: [],
    merges: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
  };
}

const part = (name: string, text: string): OpaquePart => ({ name, data: bytesToBase64(encoder.encode(text)) });

function odsBook(spec: BookSpec): SpreadsheetModel {
  const texts = spec.tables.map(tableXml);
  const prefix = HEAD + (spec.prefix ?? "");
  const suffix = (spec.suffix ?? "") + TAIL;
  const sheets = spec.tables.flatMap((table, i) => (table.deleted ? [] : [grid(table.modelName ?? decodeXmlText(table.name), texts[i]!, table.cells ?? [])]));
  for (const copy of spec.copies ?? []) {
    const origin = sheets.find((sheet) => sheet.name === copy.copiedFrom)!;
    sheets.push({ ...grid(copy.name, copy.sourceXml, copy.cells ?? origin.rows), copiedFrom: copy.copiedFrom });
  }
  return {
    dxfs: [],
    kind: "spreadsheet",
    id: "book",
    title: "Retained",
    format: "ods",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    odsContentEnvelope: { prefix, suffix },
    opaque: [
      part("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
      part("content.xml", `${spec.bom ? "\uFEFF" : ""}${prefix}${texts.join("")}${suffix}`),
      ...Object.entries(spec.parts ?? {}).map(([name, text]) => part(name, text)),
      part("styles.xml", `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}/>`),
    ],
    styles: [{}],
    defaultStyleId: 0,
  };
}

/** A frame anchored in a cell, showing embedded object `object`. */
const frame = (endCell: string, ranges: string, object = "Object 1"): string =>
  `<draw:frame table:end-cell-address="${endCell}" table:end-x="0.1in" table:end-y="0.1in" draw:z-index="0" svg:width="1in" svg:height="1in" svg:x="0in" svg:y="0in">` +
  `<draw:object draw:notify-on-update-of-ranges="${ranges}" xlink:href="./${object}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `<draw:image xlink:href="./ObjectReplacements/${object}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;

/** Table rows whose B2 holds `content` before its paragraph. */
const rowsHolding = (content: string): string =>
  `<table:table-row><table:table-cell table:number-columns-repeated="8"/></table:table-row>` +
  `<table:table-row><table:table-cell/><table:table-cell office:value-type="float" office:value="1" calcext:value-type="float">${content}<text:p>1</text:p></table:table-cell>` +
  `<table:table-cell table:number-columns-repeated="6"/></table:table-row>` +
  `<table:table-row table:number-rows-repeated="18"><table:table-cell table:number-columns-repeated="8"/></table:table-row>`;

/** The model rows of rowsHolding: B2 with its preserved content, and A7. */
const cellsHolding = (preservedXml: string): Row[] => [
  { index: 1, cells: [{ col: 1, value: 1, preservedXml }] },
  { index: 6, cells: [{ col: 0, value: 2 }] },
];

const CHART_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';

/** An embedded chart over A1:B5 of `sheet` (its local table carries a
 * print range only content.xml's rows would move). */
const chartOn = (sheet: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${CHART_NS} office:version="1.3"><office:body><office:chart>` +
  `<chart:chart chart:class="chart:bar"><chart:plot-area table:cell-range-address="${sheet}.A1:${sheet}.B5" chart:data-source-has-labels="both">` +
  `<chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="${sheet}.A2:${sheet}.A5"/></chart:axis>` +
  `<chart:series chart:values-cell-range-address="${sheet}.B2:${sheet}.B5" chart:label-cell-address="${sheet}.B1" chart:class="chart:bar">` +
  `<chart:domain table:cell-range-address="${sheet}.A2:${sheet}.A5"/><chart:data-point chart:repeated="4"/></chart:series></chart:plot-area>` +
  `<table:table table:name="local-table" table:print-ranges="${sheet}.A1:${sheet}.A2"><table:table-rows><table:table-row><table:table-cell><text:p>1</text:p></table:table-cell>` +
  `</table:table-row></table:table-rows></table:table></chart:chart></office:chart></office:body></office:document-content>`;

// --- helpers --------------------------------------------------------------------------------

function sheetNamed(model: SpreadsheetModel, name: string): Sheet {
  const sheet = model.sheets.find((each) => each.name === name);
  expect(sheet, name).toBeDefined();
  return sheet!;
}

const partOf = (model: SpreadsheetModel, name: string): OpaquePart => model.opaque.find((each) => each.name === name)!;
const textOf = (model: SpreadsheetModel, name: string): string => decoder.decode(base64ToBytes(partOf(model, name).data));
const prefixOf = (model: SpreadsheetModel): string => model.odsContentEnvelope!.prefix;
const suffixOf = (model: SpreadsheetModel): string => model.odsContentEnvelope!.suffix;
const sourceOf = (model: SpreadsheetModel, name: string): string => sheetNamed(model, name).sourceXml!;

/** One structural edit's retained rewrite, as the `retained` adjuster runs it. */
function edit(model: SpreadsheetModel, sheet: string, kind: StructuralEdit["kind"], at: number, count: number): Set<string> {
  return rewriteOdsRetainedReferences(model, structuralMapper(model, { kind, sheetId: sheetNamed(model, sheet).id, at, count }));
}

/** The changed attribute values and text nodes as [before, after] pairs. */
const spans = (before: string, after: string, removed: readonly string[] = []): [string, string][] =>
  changedSpans(before, after, removed).map((span) => [span.before, span.after]);

/** content.xml is exactly prefix + its source tables + suffix, and every
 * original's sourceXml is one of those tables. */
function expectConsistent(model: SpreadsheetModel): void {
  const content = textOf(model, "content.xml");
  const prefix = prefixOf(model);
  const suffix = suffixOf(model);
  const tables = sourceTables(content, prefix.length, suffix.length).map((table) => content.slice(table.start, table.end));
  expect(prefix + tables.join("") + suffix).toBe(content);
  for (const sheet of model.sheets) {
    if (sheet.copiedFrom === undefined && sheet.sourceXml !== undefined) {
      expect(tables).toContain(sheet.sourceXml);
    }
  }
}

const ids = (...sheets: Sheet[]): Set<string> => new Set(sheets.map((sheet) => sheet.id));

// --- the table, row by row ------------------------------------------------------------------

describe("the ods location table", () => {
  it("envelope prefix: a content validation's base cell and every bracketed reference in its condition", () => {
    const validations =
      `<table:content-validations>` +
      `<table:content-validation table:name="val1" table:condition="of:cell-content-is-in-list([$Data.$A$10:.$A$12])" table:allow-empty-cell="true" table:base-cell-address="Data.A2"/>` +
      `<table:content-validation table:name="val2" table:condition="of:cell-content() &gt; [.B3] and cell-content() &lt;&gt; &quot;[.B3]&quot;" table:base-cell-address="Data.B3">` +
      `<table:error-message table:display="true"/></table:content-validation>` +
      `<table:content-validation table:name="val3" table:condition="of:cell-content-is-in-list([$Other.$A$1:.$A$5])" table:base-cell-address="Other.A2"/>` +
      `</table:content-validations>`;
    const model = odsBook({ prefix: validations, tables: [{ name: "Data" }, { name: "Other" }] });
    const before = { prefix: prefixOf(model), suffix: suffixOf(model), sheets: model.sheets, envelope: model.odsContentEnvelope };

    const changed = edit(model, "Data", "insertRows", 1, 2);

    // The sheet-less [.B3] is on Data (val2's base cell); the string "[.B3]" and every entity stay.
    expect(spans(before.prefix, prefixOf(model))).toEqual([
      ["of:cell-content-is-in-list([$Data.$A$10:.$A$12])", "of:cell-content-is-in-list([$Data.$A$12:.$A$14])"],
      ["Data.A2", "Data.A4"],
      ["of:cell-content() &gt; [.B3] and cell-content() &lt;&gt; &quot;[.B3]&quot;", "of:cell-content() &gt; [.B5] and cell-content() &lt;&gt; &quot;[.B3]&quot;"],
      ["Data.B3", "Data.B5"],
    ]);
    expect(model.odsContentEnvelope).not.toBe(before.envelope);
    expect(suffixOf(model)).toBe(before.suffix);
    // Workbook-level text: no sheet's own text changed.
    expect(changed).toEqual(new Set());
    expect(model.sheets).toBe(before.sheets);
    expectConsistent(model);
  });

  it("envelope prefix: a label range's label and data ranges; an emptied one is removed", () => {
    const second = `<table:label-range table:label-cell-range-address="$Data.$E$3:.$E$4" table:data-cell-range-address="$Data.$F$3:.$G$4" table:orientation="row"/>`;
    const labels =
      `<table:label-ranges><table:label-range table:label-cell-range-address="$Data.$A$1:.$C$1" table:data-cell-range-address="$Data.$A$2:.$C$9" table:orientation="column"/>` +
      `${second}</table:label-ranges>`;
    const build = (): SpreadsheetModel => odsBook({ prefix: labels, tables: [{ name: "Data" }] });

    const inserted = build();
    const before = prefixOf(inserted);
    edit(inserted, "Data", "insertRows", 1, 2);
    expect(spans(before, prefixOf(inserted))).toEqual([
      ["$Data.$A$2:.$C$9", "$Data.$A$4:.$C$11"],
      ["$Data.$E$3:.$E$4", "$Data.$E$5:.$E$6"],
      ["$Data.$F$3:.$G$4", "$Data.$F$5:.$G$6"],
    ]);
    expectConsistent(inserted);

    const deleted = build();
    edit(deleted, "Data", "deleteRows", 2, 2);
    expect(spans(before, prefixOf(deleted), [second])).toEqual([["$Data.$A$2:.$C$9", "$Data.$A$2:.$C$7"]]);
    expectConsistent(deleted);
  });

  it("suffix and inside tables: named ranges' ranges and base cells, a named expression's references", () => {
    const names =
      `<table:named-expressions>` +
      `<table:named-range table:name="Block" table:base-cell-address="$Data.$A$5" table:cell-range-address="$Data.$A$5:.$C$7"/>` +
      `<table:named-range table:name="Elsewhere" table:base-cell-address="$Other.$A$1" table:cell-range-address="$Other.$A$1:.$B$2"/>` +
      `<table:named-range table:name="Unknown" table:base-cell-address="$Gone.$A$5" table:cell-range-address="$Gone.$A$5:.$C$7"/>` +
      `<table:named-expression table:name="Rate" table:base-cell-address="$Data.$B$2" table:expression="of:=[$Data.$B$4]*2+[.B4]"/>` +
      `</table:named-expressions>`;
    const local = `<table:named-expressions><table:named-range table:name="Local" table:base-cell-address="$Data.$A$3" table:cell-range-address=".$A$3:.$A$6"/></table:named-expressions>`;
    const model = odsBook({ suffix: names, tables: [{ name: "Data", trailing: local }, { name: "Other" }] });
    const before = { prefix: prefixOf(model), suffix: suffixOf(model), data: sourceOf(model, "Data"), other: sheetNamed(model, "Other") };

    const changed = edit(model, "Data", "insertRows", 1, 2);

    // A sheet no model sheet holds (Gone) is left alone; Rate's [.B4] is on its base cell's sheet.
    expect(spans(before.suffix, suffixOf(model))).toEqual([
      ["$Data.$A$5", "$Data.$A$7"],
      ["$Data.$A$5:.$C$7", "$Data.$A$7:.$C$9"],
      ["$Data.$B$2", "$Data.$B$4"],
      ["of:=[$Data.$B$4]*2+[.B4]", "of:=[$Data.$B$6]*2+[.B6]"],
    ]);
    // The sheet-local range's sheet-less entry is on its table's sheet.
    expect(spans(before.data, sourceOf(model, "Data"))).toEqual([
      ["$Data.$A$3", "$Data.$A$5"],
      [".$A$3:.$A$6", ".$A$5:.$A$8"],
    ]);
    expect(prefixOf(model)).toBe(before.prefix);
    expect(changed).toEqual(ids(sheetNamed(model, "Data")));
    expect(sheetNamed(model, "Other")).toBe(before.other);
    expectConsistent(model);
  });

  it("suffix: a database range's target and its filter's criteria range; an emptied range is removed, an emptied filter range kept", () => {
    const gone = `<table:database-range table:name="db2" table:target-range-address="Data.B1:Data.B9"/>`;
    const databases =
      `<table:database-ranges><table:database-range table:name="db1" table:target-range-address="Data.A1:Data.C9">` +
      `<table:filter table:condition-source-range-address="Data.E1:Data.E3"><table:filter-condition table:field-number="0" table:value="1" table:operator="="/></table:filter>` +
      `</table:database-range>${gone}<table:database-range table:name="db3" table:target-range-address=".A1:.C9"/></table:database-ranges>`;
    const build = (): SpreadsheetModel => odsBook({ suffix: databases, tables: [{ name: "Data" }] });
    const before = suffixOf(build());

    const inserted = build();
    edit(inserted, "Data", "insertCols", 1, 1);
    // db3's sheet-less target has no host in the envelope (no base cell): left alone.
    expect(spans(before, suffixOf(inserted))).toEqual([
      ["Data.A1:Data.C9", "Data.A1:Data.D9"],
      ["Data.E1:Data.E3", "Data.F1:Data.F3"],
      ["Data.B1:Data.B9", "Data.C1:Data.C9"],
    ]);
    expectConsistent(inserted);

    const deleted = build();
    edit(deleted, "Data", "deleteCols", 1, 1);
    expect(spans(before, suffixOf(deleted), [gone])).toEqual([
      ["Data.A1:Data.C9", "Data.A1:Data.B9"],
      ["Data.E1:Data.E3", "Data.D1:Data.D3"],
    ]);
    expectConsistent(deleted);

    // The filter's criteria column deleted: its list keeps its text — nothing changes.
    const criteria = build();
    const opaque = criteria.opaque;
    const envelope = criteria.odsContentEnvelope;
    expect(edit(criteria, "Data", "deleteCols", 4, 1)).toEqual(new Set());
    expect(criteria.opaque).toBe(opaque);
    expect(criteria.odsContentEnvelope).toBe(envelope);
  });

  it("suffix: a pilot table's target range and its source cell range (nothing else of it)", () => {
    const pivots =
      `<table:data-pilot-tables><table:data-pilot-table table:name="pivot1" table:target-range-address="Other.A10:Other.C20" table:buttons="Other.A10 Other.B10">` +
      `<table:source-cell-range table:cell-range-address="Data.A1:Data.C9"/></table:data-pilot-table></table:data-pilot-tables>`;
    const model = odsBook({ suffix: pivots, tables: [{ name: "Data" }, { name: "Other" }] });
    const before = suffixOf(model);

    edit(model, "Other", "insertRows", 0, 1);
    const afterRows = suffixOf(model);
    // table:buttons is not in the table: untouched.
    expect(spans(before, afterRows)).toEqual([["Other.A10:Other.C20", "Other.A11:Other.C21"]]);

    edit(model, "Data", "deleteCols", 1, 1);
    expect(spans(afterRows, suffixOf(model))).toEqual([["Data.A1:Data.C9", "Data.A1:Data.B9"]]);
    expectConsistent(model);
  });

  it("SPEC-128: a database range that is a sheet's autofilter is the model's — its target and its filter's criteria range are left alone; every other one still moves", () => {
    // Data's anonymous range with filter buttons is its autofilter (ods/databaseRanges.ts's claim);
    // the named one with buttons is not (the anonymous wins); Other has a named one alone, so it is Other's.
    const anonymous =
      `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Data.A1:Data.C9" table:display-filter-buttons="true">` +
      `<table:filter table:condition-source-range-address="Data.E1:Data.E3"><table:filter-condition table:field-number="0" table:value="1" table:operator="="/></table:filter></table:database-range>`;
    const named = `<table:database-range table:name="db1" table:target-range-address="Data.A12:Data.C20" table:display-filter-buttons="true"/>`;
    const other = `<table:database-range table:name="db2" table:target-range-address="Other.A1:Other.B5" table:display-filter-buttons="true"/>`;
    const plain = `<table:database-range table:name="db3" table:target-range-address="Data.B1:Data.B9"/>`;
    const build = (): SpreadsheetModel =>
      odsBook({ suffix: `<table:database-ranges>${anonymous}${named}${other}${plain}</table:database-ranges>`, tables: [{ name: "Data" }, { name: "Other" }] });
    const before = suffixOf(build());

    const inserted = build();
    edit(inserted, "Data", "insertCols", 0, 1);
    expect(spans(before, suffixOf(inserted))).toEqual([
      ["Data.A12:Data.C20", "Data.B12:Data.D20"],
      ["Data.B1:Data.B9", "Data.C1:Data.C9"],
    ]);
    expect(suffixOf(inserted)).toContain(anonymous);
    expectConsistent(inserted);

    // Deleting every column of the owned range would empty it: it is not removed here either.
    const deleted = build();
    edit(deleted, "Data", "deleteCols", 0, 3);
    expect(suffixOf(deleted)).toContain(anonymous);
    expect(suffixOf(deleted)).not.toContain(named);
    expect(suffixOf(deleted)).not.toContain(plain);
    expectConsistent(deleted);

    const rows = build();
    expect(edit(rows, "Other", "insertRows", 0, 2)).toEqual(new Set());
    expect(suffixOf(rows)).toBe(before);
  });

  it("suffix: a consolidation's source ranges and target cell, a deleted target becoming the band's start", () => {
    const consolidation = `<table:consolidation table:function="sum" table:source-cell-range-addresses="Data.A1:Data.B4 Other.A1:Other.B4" table:target-cell-address="Data.H6"/>`;
    const build = (): SpreadsheetModel => odsBook({ suffix: consolidation, tables: [{ name: "Data" }, { name: "Other" }] });
    const before = suffixOf(build());

    const inserted = build();
    edit(inserted, "Data", "insertRows", 1, 2);
    expect(spans(before, suffixOf(inserted))).toEqual([
      ["Data.A1:Data.B4 Other.A1:Other.B4", "Data.A1:Data.B6 Other.A1:Other.B4"],
      ["Data.H6", "Data.H8"],
    ]);

    // Rows 5–6 deleted: H6 is gone, so the anchor goes to where the band now starts (H5).
    const deleted = build();
    edit(deleted, "Data", "deleteRows", 4, 2);
    expect(spans(before, suffixOf(deleted))).toEqual([["Data.H6", "Data.H5"]]);
    expectConsistent(deleted);
  });

  it("table start tag: print ranges — moved parts upper-cased, the rest verbatim, a gone entry dropped, an emptied attribute removed", () => {
    const build = (): SpreadsheetModel =>
      odsBook({
        tables: [
          { name: "Data", attrs: ` table:print-ranges="$Data.$a$1:.$d$20  $Data.$F$3:.$F$4"` },
          { name: "Other", attrs: ` table:print-ranges="$Other.$B$3:.$C$4"` },
        ],
      });
    const before = { data: sourceOf(build(), "Data"), other: sourceOf(build(), "Other") };

    const inserted = build();
    expect(edit(inserted, "Data", "insertRows", 1, 2)).toEqual(ids(sheetNamed(inserted, "Data")));
    expect(spans(before.data, sourceOf(inserted, "Data"))).toEqual([["$Data.$a$1:.$d$20  $Data.$F$3:.$F$4", "$Data.$a$1:.$D$22  $Data.$F$5:.$F$6"]]);
    expectConsistent(inserted);

    const deleted = build();
    edit(deleted, "Data", "deleteRows", 2, 2);
    expect(spans(before.data, sourceOf(deleted, "Data"))).toEqual([["$Data.$a$1:.$d$20  $Data.$F$3:.$F$4", "$Data.$a$1:.$D$18"]]);
    expectConsistent(deleted);

    const emptied = build();
    edit(emptied, "Other", "deleteRows", 2, 2);
    expect(spans(before.other, sourceOf(emptied, "Other"), [` table:print-ranges="$Other.$B$3:.$C$4"`])).toEqual([]);
    expect(sourceOf(emptied, "Other")).not.toContain("print-ranges");
    expectConsistent(emptied);
  });

  it("inside tables: conditional formats' targets, conditions' values and base cells, formatting entries' values", () => {
    const formats =
      `<calcext:conditional-formats>` +
      `<calcext:conditional-format calcext:target-range-address="Data.B2:Data.B9">` +
      `<calcext:condition calcext:apply-style-name="Bad" calcext:value="formula-is([.B2]&gt;[$Other.$B$1])" calcext:base-cell-address="Data.B2"/>` +
      `</calcext:conditional-format>` +
      `<calcext:conditional-format calcext:target-range-address="Data.C2:Data.C9 Data.E2:Data.E3"><calcext:color-scale>` +
      `<calcext:formatting-entry calcext:value="[.C2]" calcext:type="formula" calcext:color="#ff0000"/>` +
      `<calcext:formatting-entry calcext:value="0" calcext:type="maximum" calcext:color="#00ff00"/>` +
      `</calcext:color-scale></calcext:conditional-format>` +
      `</calcext:conditional-formats>`;
    const model = odsBook({ tables: [{ name: "Data", trailing: formats }, { name: "Other" }] });
    const before = sourceOf(model, "Data");

    expect(edit(model, "Data", "insertRows", 1, 2)).toEqual(ids(sheetNamed(model, "Data")));
    expect(spans(before, sourceOf(model, "Data"))).toEqual([
      ["Data.B2:Data.B9", "Data.B4:Data.B11"],
      ["formula-is([.B2]&gt;[$Other.$B$1])", "formula-is([.B4]&gt;[$Other.$B$1])"],
      ["Data.B2", "Data.B4"],
      ["Data.C2:Data.C9 Data.E2:Data.E3", "Data.C4:Data.C11 Data.E4:Data.E5"],
      ["[.C2]", "[.C4]"],
    ]);
    expectConsistent(model);
  });

  it("inside tables: an emptied conditional format is removed, and so is an emptied container", () => {
    const gone = `<calcext:conditional-format calcext:target-range-address="Data.B3:Data.B4"><calcext:condition calcext:value="[.B3]&gt;1" calcext:base-cell-address="Data.B3"/></calcext:conditional-format>`;
    const kept = `<calcext:conditional-format calcext:target-range-address="Data.C3:Data.D9"><calcext:condition calcext:value="[.C3]&gt;1" calcext:base-cell-address="Data.C3"/></calcext:conditional-format>`;

    const mixed = odsBook({ tables: [{ name: "Data", trailing: `<calcext:conditional-formats>${gone}${kept}</calcext:conditional-formats>` }] });
    const mixedBefore = sourceOf(mixed, "Data");
    edit(mixed, "Data", "deleteRows", 2, 2);
    // [.C3] lay in the band: #REF!; the base cell C3 goes where the band starts — C3 again, so it keeps its bytes.
    expect(spans(mixedBefore, sourceOf(mixed, "Data"), [gone])).toEqual([
      ["Data.C3:Data.D9", "Data.C3:Data.D7"],
      ["[.C3]&gt;1", "[.#REF!]&gt;1"],
    ]);
    expectConsistent(mixed);

    const container = `<calcext:conditional-formats>${gone}</calcext:conditional-formats>`;
    const emptied = odsBook({ tables: [{ name: "Data", trailing: container }] });
    const emptiedBefore = sourceOf(emptied, "Data");
    edit(emptied, "Data", "deleteRows", 2, 2);
    expect(spans(emptiedBefore, sourceOf(emptied, "Data"), [container])).toEqual([]);
    expect(sourceOf(emptied, "Data")).not.toContain("calcext:");
    expectConsistent(emptied);
  });

  it("inside tables and Cell.preservedXml: shapes' end cells and an object's notify ranges, copy-on-write", () => {
    const anchored = frame("Data.D6", "Data.A1:Data.B5 Other.A1:Other.A3");
    const shapes = `<table:shapes><draw:custom-shape table:end-cell-address="Data.C3" table:end-x="0in" table:end-y="0in" draw:z-index="1" svg:width="1in" svg:height="1in" svg:x="0in" svg:y="0in"/></table:shapes>`;
    const build = (): SpreadsheetModel =>
      odsBook({ tables: [{ name: "Data", leading: shapes, rows: rowsHolding(anchored), cells: cellsHolding(anchored) }, { name: "Other" }] });

    const model = build();
    const before = sheetNamed(model, "Data");
    const changed = edit(model, "Data", "insertRows", 1, 2);
    const after = sheetNamed(model, "Data");
    expect(spans(before.sourceXml!, after.sourceXml!)).toEqual([
      ["Data.C3", "Data.C5"],
      ["Data.D6", "Data.D8"],
      ["Data.A1:Data.B5 Other.A1:Other.A3", "Data.A1:Data.B7 Other.A1:Other.A3"],
    ]);
    expect(spans(before.rows[0]!.cells[0]!.preservedXml!, after.rows[0]!.cells[0]!.preservedXml!)).toEqual([
      ["Data.D6", "Data.D8"],
      ["Data.A1:Data.B5 Other.A1:Other.A3", "Data.A1:Data.B7 Other.A1:Other.A3"],
    ]);
    // New rows array, row and cell for the change; the untouched row keeps its identity; nothing mutated.
    expect(after).not.toBe(before);
    expect(after.rows).not.toBe(before.rows);
    expect(after.rows[0]).not.toBe(before.rows[0]);
    expect(after.rows[1]).toBe(before.rows[1]);
    expect(before.rows[0]!.cells[0]!.preservedXml).toBe(anchored);
    expect(changed).toEqual(ids(after));
    expect(sheetNamed(model, "Other")).toBe(model.sheets[1]);
    expectConsistent(model);

    // An edit on Other moves only Other's entry inside Data's text — Data's own text changed.
    const other = build();
    expect(edit(other, "Other", "insertRows", 0, 1)).toEqual(ids(sheetNamed(other, "Data")));
    expect(spans(anchored, sheetNamed(other, "Data").rows[0]!.cells[0]!.preservedXml!)).toEqual([
      ["Data.A1:Data.B5 Other.A1:Other.A3", "Data.A1:Data.B5 Other.A2:Other.A4"],
    ]);
    expectConsistent(other);
  });

  it("inside tables: a scenario's ranges; an emptied list keeps its text", () => {
    const scenario =
      `<table:scenario table:scenario-ranges="Data.B4:Data.C6" table:is-active="true" table:display-border="true" table:border-color="#c0c0c0" ` +
      `table:copy-back="true" table:copy-styles="true" table:copy-formulas="true" table:comment=""/>`;
    const build = (): SpreadsheetModel => odsBook({ tables: [{ name: "Data", leading: scenario }] });

    const inserted = build();
    const before = sourceOf(inserted, "Data");
    edit(inserted, "Data", "insertCols", 0, 2);
    expect(spans(before, sourceOf(inserted, "Data"))).toEqual([["Data.B4:Data.C6", "Data.D4:Data.E6"]]);
    expectConsistent(inserted);

    const deleted = build();
    const { sheets, opaque } = deleted;
    expect(edit(deleted, "Data", "deleteRows", 3, 3)).toEqual(new Set());
    expect(deleted.sheets).toBe(sheets);
    expect(deleted.opaque).toBe(opaque);
  });

  it("Object N/content.xml: the chart's plot area, series values and label, domain and categories", () => {
    // Other shows the chart; the chart reads Data.
    const shown = frame("Other.D6", "Other.A1:Other.A3");
    const model = odsBook({
      tables: [{ name: "Data" }, { name: "Other", rows: rowsHolding(shown), cells: cellsHolding(shown) }],
      parts: { "Object 1/content.xml": chartOn("Data"), "Object 1/styles.xml": `<?xml version="1.0" encoding="UTF-8"?><office:document-styles/>` },
    });
    const before = { chart: textOf(model, "Object 1/content.xml"), opaque: model.opaque, sheets: model.sheets, styles: partOf(model, "Object 1/styles.xml") };
    const content = partOf(model, "content.xml");

    const changed = edit(model, "Data", "insertCols", 0, 1);

    // The local table's print range is not a chart location.
    expect(spans(before.chart, textOf(model, "Object 1/content.xml"))).toEqual([
      ["Data.A1:Data.B5", "Data.B1:Data.C5"],
      ["Data.A2:Data.A5", "Data.B2:Data.B5"],
      ["Data.B2:Data.B5", "Data.C2:Data.C5"],
      ["Data.B1", "Data.C1"],
      ["Data.A2:Data.A5", "Data.B2:Data.B5"],
    ]);
    // The sheet whose table shows the chart is reported; its own text is unchanged.
    expect(changed).toEqual(ids(sheetNamed(model, "Other")));
    expect(model.sheets).toBe(before.sheets);
    expect(model.opaque).not.toBe(before.opaque);
    expect(partOf(model, "content.xml")).toBe(content);
    expect(partOf(model, "Object 1/styles.xml")).toBe(before.styles);
  });
});

// --- rules across the table -----------------------------------------------------------------

describe("ods retained references", () => {
  it("leaves 3-D, external and #REF! references alone and turns a gone area into [<sheet as written>.#REF!]", () => {
    const names =
      `<table:named-expressions>` +
      `<table:named-range table:name="Span" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$A$5:$Other.$B$5"/>` +
      `<table:named-expression table:name="Ext" table:base-cell-address="$Data.$A$1" table:expression="of:=['file:///tmp/x.ods'#$Data.A5]+[$Data.A5]+[.#REF!]+[$Data.A5:$Other.A5]"/>` +
      `<table:named-expression table:name="Lost" table:base-cell-address="$Data.$B$2" table:expression="of:=[$Data.A3]+SUM([.A3:.A4])*2"/>` +
      `</table:named-expressions>`;
    const build = (): SpreadsheetModel => odsBook({ suffix: names, tables: [{ name: "Data" }, { name: "Other" }] });
    const before = suffixOf(build());

    const inserted = build();
    edit(inserted, "Data", "insertRows", 1, 2);
    expect(spans(before, suffixOf(inserted))).toEqual([
      ["of:=['file:///tmp/x.ods'#$Data.A5]+[$Data.A5]+[.#REF!]+[$Data.A5:$Other.A5]", "of:=['file:///tmp/x.ods'#$Data.A5]+[$Data.A7]+[.#REF!]+[$Data.A5:$Other.A5]"],
      ["$Data.$B$2", "$Data.$B$4"],
      ["of:=[$Data.A3]+SUM([.A3:.A4])*2", "of:=[$Data.A5]+SUM([.A5:.A6])*2"],
    ]);

    const deleted = build();
    edit(deleted, "Data", "deleteRows", 2, 2);
    expect(spans(before, suffixOf(deleted))).toEqual([
      ["of:=['file:///tmp/x.ods'#$Data.A5]+[$Data.A5]+[.#REF!]+[$Data.A5:$Other.A5]", "of:=['file:///tmp/x.ods'#$Data.A5]+[$Data.A3]+[.#REF!]+[$Data.A5:$Other.A5]"],
      ["of:=[$Data.A3]+SUM([.A3:.A4])*2", "of:=[$Data.#REF!]+SUM([.#REF!])*2"],
    ]);
    expectConsistent(deleted);
  });

  it("never touches table:tracked-changes", () => {
    const tracked =
      `<table:tracked-changes table:track-changes="true">` +
      `<table:cell-content-change table:id="ct1"><table:cell-address table:column="1" table:row="8" table:table="0"/>` +
      `<office:change-info><dc:creator>Ann</dc:creator><dc:date>2026-01-01T00:00:00</dc:date></office:change-info>` +
      `<table:previous><table:change-track-table-cell office:value-type="float" office:value="1"><draw:frame table:end-cell-address="Data.B9"/><text:p>1</text:p></table:change-track-table-cell></table:previous>` +
      `</table:cell-content-change>` +
      `<table:deletion table:type="row" table:position="3" table:id="ct2"><office:change-info><dc:creator>Ann</dc:creator><dc:date>2026-01-01T00:00:00</dc:date></office:change-info>` +
      `<table:content-validation table:name="old" table:condition="of:[.A9]&gt;0" table:base-cell-address="Data.A9"/></table:deletion>` +
      `</table:tracked-changes>`;
    const validations = `<table:content-validations><table:content-validation table:name="val1" table:condition="of:[.A9]&gt;0" table:base-cell-address="Data.A9"/></table:content-validations>`;
    const model = odsBook({ prefix: tracked + validations, tables: [{ name: "Data" }] });
    const before = prefixOf(model);

    edit(model, "Data", "insertRows", 1, 2);
    expect(spans(before, prefixOf(model))).toEqual([
      ["of:[.A9]&gt;0", "of:[.A11]&gt;0"],
      ["Data.A9", "Data.A11"],
    ]);
    expect(prefixOf(model)).toContain(tracked);
    expectConsistent(model);
  });

  it("resolves source-time names: a renamed sheet's references move and keep the name as written", () => {
    const names = `<table:named-expressions><table:named-range table:name="Block" table:base-cell-address="$Data.$A$5" table:cell-range-address="$Data.$A$5:.$C$7"/></table:named-expressions>`;
    const local = `<table:named-expressions><table:named-range table:name="Local" table:base-cell-address="$Data.$A$3" table:cell-range-address="$Data.$A$3:.$A$6"/></table:named-expressions>`;
    const model = odsBook({ suffix: names, tables: [{ name: "Data", modelName: "Ledger", trailing: local }] });
    const before = { suffix: suffixOf(model), source: sourceOf(model, "Ledger") };

    expect(edit(model, "Ledger", "insertRows", 1, 2)).toEqual(ids(sheetNamed(model, "Ledger")));
    expect(spans(before.suffix, suffixOf(model))).toEqual([
      ["$Data.$A$5", "$Data.$A$7"],
      ["$Data.$A$5:.$C$7", "$Data.$A$7:.$C$9"],
    ]);
    expect(spans(before.source, sourceOf(model, "Ledger"))).toEqual([
      ["$Data.$A$3", "$Data.$A$5"],
      ["$Data.$A$3:.$A$6", "$Data.$A$5:.$A$8"],
    ]);
    expect(sourceOf(model, "Ledger")).toContain(`table:name="Data"`);
    expectConsistent(model);
  });

  it("keeps entities in sheet names byte-for-byte", () => {
    const names =
      `<table:named-expressions>` +
      `<table:named-range table:name="qa" table:base-cell-address="$Q&amp;A.$A$2" table:cell-range-address="$Q&amp;A.$A$2:.$A$3"/>` +
      `<table:named-range table:name="tom" table:base-cell-address="$&apos;Tom&apos;&apos;s&apos;.$A$2" table:cell-range-address="$&apos;Tom&apos;&apos;s&apos;.A2 $&apos;Tom&apos;&apos;s&apos;.C2"/>` +
      `</table:named-expressions>`;
    const model = odsBook({ suffix: names, tables: [{ name: "Q&amp;A" }, { name: "Tom&apos;s" }] });
    const before = suffixOf(model);

    edit(model, "Q&A", "insertRows", 0, 1);
    const afterQa = suffixOf(model);
    expect(spans(before, afterQa)).toEqual([
      ["$Q&amp;A.$A$2", "$Q&amp;A.$A$3"],
      ["$Q&amp;A.$A$2:.$A$3", "$Q&amp;A.$A$3:.$A$4"],
    ]);

    edit(model, "Tom's", "insertCols", 1, 1);
    expect(spans(afterQa, suffixOf(model))).toEqual([
      ["$&apos;Tom&apos;&apos;s&apos;.A2 $&apos;Tom&apos;&apos;s&apos;.C2", "$&apos;Tom&apos;&apos;s&apos;.A2 $&apos;Tom&apos;&apos;s&apos;.D2"],
    ]);
    expectConsistent(model);
  });

  it("moves a whole column only along columns, a whole row only along rows, and keeps a reversed range's orientation", () => {
    const names =
      `<table:named-expressions>` +
      `<table:named-range table:name="cols" table:base-cell-address="$Data.$B$1" table:cell-range-address="$Data.$B:.$C"/>` +
      `<table:named-range table:name="rows" table:base-cell-address="$Data.$A$3" table:cell-range-address="$Data.$3:.$5"/>` +
      `<table:named-range table:name="rev" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$C$9:.$A$2"/>` +
      `</table:named-expressions>`;
    const model = odsBook({ suffix: names, tables: [{ name: "Data" }] });
    const before = suffixOf(model);

    edit(model, "Data", "insertRows", 1, 2);
    const afterRows = suffixOf(model);
    expect(spans(before, afterRows)).toEqual([
      ["$Data.$A$3", "$Data.$A$5"],
      ["$Data.$3:.$5", "$Data.$5:.$7"],
      ["$Data.$C$9:.$A$2", "$Data.$C$11:.$A$4"],
    ]);

    edit(model, "Data", "insertCols", 1, 1);
    expect(spans(afterRows, suffixOf(model))).toEqual([
      ["$Data.$B$1", "$Data.$C$1"],
      ["$Data.$B:.$C", "$Data.$C:.$D"],
      ["$Data.$C$11:.$A$4", "$Data.$D$11:.$A$4"],
    ]);
    expectConsistent(model);
  });

  it("adjusts a duplicated sheet's own sourceXml, rows and chart under its copy override — and not its origin's", () => {
    const originFrame = frame("Data.D6", "Data.A1:Data.B5");
    const copyFrame = frame("Data.D6", "Data.A1:Data.B5", "Object 2");
    const build = (): SpreadsheetModel =>
      odsBook({
        tables: [{ name: "Data", rows: rowsHolding(originFrame), cells: cellsHolding(originFrame) }],
        copies: [{ name: "Data (2)", copiedFrom: "Data", sourceXml: tableXml({ name: "Data", rows: rowsHolding(copyFrame) }) }],
        parts: { "Object 1/content.xml": chartOn("Data"), "Object 2/content.xml": chartOn("Data") },
      });

    // The copy edited: its own text and its own chart move; the origin and content.xml do not.
    const onCopy = build();
    const before = {
      origin: onCopy.sheets[0]!,
      copy: onCopy.sheets[1]!,
      content: partOf(onCopy, "content.xml"),
      object1: partOf(onCopy, "Object 1/content.xml"),
      envelope: onCopy.odsContentEnvelope,
    };
    const chart = textOf(onCopy, "Object 2/content.xml");
    expect(before.copy.rows).toBe(before.origin.rows);

    expect(edit(onCopy, "Data (2)", "insertRows", 1, 2)).toEqual(ids(before.copy));
    const copy = sheetNamed(onCopy, "Data (2)");
    expect(spans(before.copy.sourceXml!, copy.sourceXml!)).toEqual([
      ["Data.D6", "Data.D8"],
      ["Data.A1:Data.B5", "Data.A1:Data.B7"],
    ]);
    expect(spans(originFrame, copy.rows[0]!.cells[0]!.preservedXml!)).toEqual([
      ["Data.D6", "Data.D8"],
      ["Data.A1:Data.B5", "Data.A1:Data.B7"],
    ]);
    // The series label B1 lies above the insert: it keeps its bytes.
    expect(spans(chart, textOf(onCopy, "Object 2/content.xml"))).toEqual([
      ["Data.A1:Data.B5", "Data.A1:Data.B7"],
      ["Data.A2:Data.A5", "Data.A4:Data.A7"],
      ["Data.B2:Data.B5", "Data.B4:Data.B7"],
      ["Data.A2:Data.A5", "Data.A4:Data.A7"],
    ]);
    expect(sheetNamed(onCopy, "Data")).toBe(before.origin);
    expect(partOf(onCopy, "content.xml")).toBe(before.content);
    expect(partOf(onCopy, "Object 1/content.xml")).toBe(before.object1);
    expect(onCopy.odsContentEnvelope).toBe(before.envelope);

    // The origin edited: the copy keeps every byte and its identity.
    const onOrigin = build();
    const copyBefore = onOrigin.sheets[1]!;
    const object2 = partOf(onOrigin, "Object 2/content.xml");
    const sourceBefore = sourceOf(onOrigin, "Data");
    expect(edit(onOrigin, "Data", "insertRows", 1, 2)).toEqual(ids(onOrigin.sheets[0]!));
    expect(spans(sourceBefore, sourceOf(onOrigin, "Data"))).toEqual([
      ["Data.D6", "Data.D8"],
      ["Data.A1:Data.B5", "Data.A1:Data.B7"],
    ]);
    expect(onOrigin.sheets[1]).toBe(copyBefore);
    expect(partOf(onOrigin, "Object 2/content.xml")).toBe(object2);
    expect(textOf(onOrigin, "Object 1/content.xml")).toContain(`table:cell-range-address="Data.A1:Data.B7"`);
    expectConsistent(onOrigin);
  });

  it("keeps every object's identity when nothing moves", () => {
    const names = `<table:named-expressions><table:named-range table:name="Block" table:base-cell-address="$Data.$A$5" table:cell-range-address="$Data.$A$5:.$C$7"/></table:named-expressions>`;
    const anchored = frame("Data.D6", "Data.A1:Data.B5");
    const model = odsBook({
      suffix: names,
      tables: [{ name: "Data", rows: rowsHolding(anchored), cells: cellsHolding(anchored) }, { name: "Other" }],
      parts: { "Object 1/content.xml": chartOn("Data") },
    });
    const before = { sheets: model.sheets, opaque: model.opaque, envelope: model.odsContentEnvelope, data: model.sheets[0]! };

    // An insert below everything, and an edit on a sheet nothing names.
    expect(edit(model, "Data", "insertRows", 30, 5)).toEqual(new Set());
    expect(edit(model, "Other", "deleteCols", 0, 3)).toEqual(new Set());
    expect(model.sheets).toBe(before.sheets);
    expect(model.sheets[0]).toBe(before.data);
    expect(model.opaque).toBe(before.opaque);
    expect(model.odsContentEnvelope).toBe(before.envelope);
  });

  it("writes content.xml back with its byte-order mark", () => {
    const names = `<table:named-expressions><table:named-range table:name="Block" table:base-cell-address="$Data.$A$5" table:cell-range-address="$Data.$A$5:.$C$7"/></table:named-expressions>`;
    const model = odsBook({ suffix: names, tables: [{ name: "Data" }], bom: true });
    edit(model, "Data", "insertRows", 0, 1);
    const bytes = base64ToBytes(partOf(model, "content.xml").data);
    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(suffixOf(model)).toContain(`table:cell-range-address="$Data.$A$6:.$C$8"`);
    expectConsistent(model);
  });

  it("adjusts a deleted sheet's table still in content.xml: its references to other sheets move, its sheet-less ones have no host", () => {
    const local =
      `<table:named-expressions><table:named-range table:name="Own" table:base-cell-address=".$A$3" table:cell-range-address=".$A$3:.$A$6"/>` +
      `<table:named-range table:name="Theirs" table:base-cell-address="$Data.$A$3" table:cell-range-address="$Data.$A$3:.$A$4"/></table:named-expressions>`;
    const model = odsBook({ tables: [{ name: "Data" }, { name: "Old", trailing: local, deleted: true }] });
    const content = textOf(model, "content.xml");
    const envelope = model.odsContentEnvelope;

    expect(edit(model, "Data", "insertRows", 1, 2)).toEqual(new Set());
    expect(spans(content, textOf(model, "content.xml"))).toEqual([
      ["$Data.$A$3", "$Data.$A$5"],
      ["$Data.$A$3:.$A$4", "$Data.$A$5:.$A$6"],
    ]);
    expect(model.odsContentEnvelope).toBe(envelope);
    expectConsistent(model);
  });

  it("leaves content.xml and the envelope alone when they do not agree", () => {
    const names = `<table:named-expressions><table:named-range table:name="Block" table:base-cell-address="$Data.$A$5" table:cell-range-address="$Data.$A$5:.$C$7"/></table:named-expressions>`;
    const local = `<table:named-expressions><table:named-range table:name="Local" table:base-cell-address="$Data.$A$3" table:cell-range-address="$Data.$A$3:.$A$6"/></table:named-expressions>`;
    const model = odsBook({ suffix: names, tables: [{ name: "Data", trailing: local }] });
    model.odsContentEnvelope = { prefix: prefixOf(model), suffix: `${suffixOf(model)} ` };
    const { opaque, odsContentEnvelope } = model;

    // The sheet's own text still moves.
    expect(edit(model, "Data", "insertRows", 1, 2)).toEqual(ids(model.sheets[0]!));
    expect(model.opaque).toBe(opaque);
    expect(model.odsContentEnvelope).toBe(odsContentEnvelope);
    expect(sourceOf(model, "Data")).toContain(`table:cell-range-address="$Data.$A$5:.$A$8"`);
  });

  it("never scans text that names no location attribute", () => {
    const names = `<table:named-expressions><table:named-range table:name="Block" table:base-cell-address="$Data.$A$5" table:cell-range-address="$Data.$A$5:.$C$7"/></table:named-expressions>`;
    const plain = `<office:annotation><dc:creator>Ann</dc:creator><text:p>note</text:p></office:annotation>`;
    const model = odsBook({
      suffix: names,
      tables: [{ name: "Data", rows: rowsHolding(plain), cells: cellsHolding(plain) }, { name: "Other" }],
      parts: { "Object 1/content.xml": `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${CHART_NS}/>` },
    });
    const scans = vi.mocked(scanXml);
    scans.mockClear();
    edit(model, "Data", "insertRows", 1, 2);
    // Only the envelope (it holds Block); no table, sourceXml, preservedXml or object.
    expect(scans).toHaveBeenCalledTimes(1);
    expect(suffixOf(model)).toContain(`table:cell-range-address="$Data.$A$7:.$C$9"`);
  });

  it("is what rewriteRetainedReferences runs for an ods workbook, marking the sheets whose text changed dirty", () => {
    const local = `<table:named-expressions><table:named-range table:name="Local" table:base-cell-address="$Data.$A$3" table:cell-range-address=".$A$3:.$A$6"/></table:named-expressions>`;
    const model = odsBook({ tables: [{ name: "Data", trailing: local }, { name: "Other" }] });
    const other = model.sheets[1]!;
    rewriteRetainedReferences(model, structuralMapper(model, { kind: "deleteRows", sheetId: model.sheets[0]!.id, at: 0, count: 1 }));
    expect(model.sheets[0]!.dirty).toBe(true);
    expect(model.sheets[0]!.sourceXml).toContain(`table:cell-range-address=".$A$2:.$A$5"`);
    expect(model.sheets[1]).toBe(other);
    expectConsistent(model);
  });
});
