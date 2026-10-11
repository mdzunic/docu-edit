// SPEC-128 "xlsx wire mapping" (AC-22's xlsx half, AC-33, AC-36's xlsx half,
// AC-41): every row of the wire table both ways — read → model, and model →
// xml → read again equal — over synthetic parts; a modelled element's
// sourceXml re-emitted verbatim inside a rebuilt (dirty) sheet and serialised
// when it does not fit the part; a source element the model no longer has
// removed; `hidden="1"` for either kind of hidden row and `sheetPr@filterMode`
// exactly when a row is filter-hidden; the `_xlnm._FilterDatabase`
// reconciliation (added, moved, removed, through reorder/delete/duplicate,
// from scratch, untouched); the read classification of hidden rows; and
// spec128-filter.xlsx read and written.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { parseRange, type GridRange } from "../address";
import { buildSpec128FilterXlsxBytes, SPEC128_FILTER_XLSX } from "../fixtures/build";
import type { AutoFilter, FilterCriteria } from "../filter";
import { CellBudget } from "../parsed";
import type { Row, Sheet, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook, deleteSheet, duplicateSheet, moveSheet, renameSheet } from "../sheets";
import type { SortState } from "../sort";
import { filterDatabaseNames, filterDatabaseText, fitsPart, readAutoFilterXml, readSheetFilterElements, readSortStateXml, top10Threshold } from "./autoFilter";
import { readWorksheet, serializeWorksheet, type WorksheetContent, type XlsxSheetContext } from "./sheetXml";
import { writeXlsx } from "./writer";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const ROOT = `<worksheet xmlns="${MAIN}" xmlns:r="${R}" xmlns:xr="${XR}">`;
const UID = `{00000000-0001-0000-0000-000000000000}`;
const encoder = new TextEncoder();
const decoder = new TextDecoder();

const range = (text: string): GridRange => parseRange(text)!;

/** A worksheet part around `body` (the root's children). */
const worksheet = (body: string, root = ROOT): string => `${DECL}${root}${body}</worksheet>`;

/** Reads a worksheet whose root holds `body`. */
function read(body: string, root = ROOT): WorksheetContent {
  const content = readWorksheet(worksheet(body, root), "S", { sharedStrings: [], dateSystem: "1900", budget: new CellBudget(), mdw: 7 });
  if (content === null) {
    throw new Error("the cell ceiling stopped the read");
  }
  return content;
}

/** A grid sheet with SPEC-117's defaults and `fields`. */
function sheetOf(fields: Partial<Sheet> = {}): Sheet {
  return { id: "s", name: "S", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

/** A save context: a string cell's index is 100 + its order, a style id its own xf. */
function stubContext(): XlsxSheetContext {
  const asked: string[] = [];
  return {
    active: false,
    theme: undefined,
    mdw: 7,
    strings: { indexFor: (cell) => 100 + (asked.includes(String(cell.value)) ? asked.indexOf(String(cell.value)) : asked.push(String(cell.value)) - 1) },
    styles: { xfFor: (id) => id },
  };
}

/** What serializeWorksheet writes for `sheet` into a part holding `body`. */
const written = (sheet: Sheet, body = `<sheetData/>`, root = ROOT): string => serializeWorksheet(sheet, worksheet(body, root), stubContext());

/** The part's root-level `<local>` element, as written. */
function elementIn(xml: string, local: string): string | undefined {
  const match = new RegExp(`<((?:[\\w.-]+:)?)${local}\\b[^>]*?(?:/>|>[\\s\\S]*?</\\1${local}>)`).exec(xml);
  return match?.[0];
}

/** A model value with its sourceXml fields dropped — what a write → read
 * round trip compares. */
function withoutSource<T extends AutoFilter | SortState | undefined>(value: T): T {
  if (value === undefined) {
    return value;
  }
  const { sourceXml: _source, ...rest } = value;
  if ("sort" in rest && rest.sort !== undefined) {
    const { sourceXml: _sortSource, ...sort } = rest.sort;
    return { ...rest, sort } as T;
  }
  return rest as T;
}

/** Numbered rows with a text cell in column A and a number in column B. */
function rowsOf(...cells: Array<[string | null, number | null, Partial<Row>?]>): Row[] {
  return cells.map(([text, amount, extra], i) => ({
    index: i + 1,
    cells: [...(text === null ? [] : [{ col: 0, value: text }]), ...(amount === null ? [] : [{ col: 1, value: amount }])],
    ...extra,
  }));
}

// --- reading ---------------------------------------------------------------------------------

describe("reading: the xlsx wire table → the model", () => {
  it("autoFilter@ref → range; the whole element, xr:uid and all, → sourceXml", () => {
    const xml = `<autoFilter ref="A1:D13" xr:uid="${UID}"/>`;
    expect(read(`<sheetData/>${xml}`).autoFilter).toEqual({ range: range("A1:D13"), columns: [], sourceXml: xml });
    const withChildren = `<autoFilter ref="B3:C9" xr:uid="${UID}"><filterColumn colId="1" hiddenButton="1"/><extLst><ext uri="{X}"/></extLst></autoFilter>`;
    expect(read(`<sheetData/>${withChildren}`).autoFilter).toEqual({ range: range("B3:C9"), columns: [{ colId: 1, hideButton: true }], sourceXml: withChildren });
  });

  it("filterColumn@colId; hiddenButton=\"1\" or showButton=\"0\" → hideButton; columns ascending by colId, the first per colId kept", () => {
    const filter = read(
      `<sheetData/><autoFilter ref="A1:F9"><filterColumn colId="4" showButton="0"/><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn>` +
        `<filterColumn colId="0" hiddenButton="1" showButton="1"/><filterColumn colId="2"><filters><filter val="y"/></filters></filterColumn>` +
        `<filterColumn colId="5" showButton="1"/><filterColumn colId="nope"/></autoFilter>`,
    ).autoFilter!;
    expect(filter.columns).toEqual([
      { colId: 0, hideButton: true },
      { colId: 2, criteria: { kind: "values", values: ["x"], blanks: false, dateGroups: [] } },
      { colId: 4, hideButton: true },
      { colId: 5 },
    ]);
  });

  it("filters@blank, filter@val and dateGroupItem → values (a dateGroupItem without year or with an unknown grouping skipped)", () => {
    const criteria = read(
      `<sheetData/><autoFilter ref="A1:A9"><filterColumn colId="0"><filters blank="1" calendarType="gregorian"><filter val="East"/><filter val="A &amp; B"/>` +
        `<filter/><dateGroupItem year="2026" month="3" dateTimeGrouping="month"/><dateGroupItem year="2025" month="12" day="31" hour="23" minute="59" second="58" dateTimeGrouping="second"/>` +
        `<dateGroupItem month="1" dateTimeGrouping="month"/><dateGroupItem year="2024" dateTimeGrouping="week"/></filters></filterColumn></autoFilter>`,
    ).autoFilter!.columns[0]!.criteria;
    expect(criteria).toEqual({
      kind: "values",
      values: ["East", "A & B"],
      blanks: true,
      dateGroups: [
        { grouping: "month", year: 2026, month: 3 },
        { grouping: "second", year: 2025, month: 12, day: 31, hour: 23, minute: 59, second: 58 },
      ],
    });
    expect(read(`<sheetData/><autoFilter ref="A1:A9"><filterColumn colId="0"><filters/></filterColumn></autoFilter>`).autoFilter!.columns[0]!.criteria).toEqual({
      kind: "values",
      values: [],
      blanks: false,
      dateGroups: [],
    });
  });

  it("customFilters@and=\"1\" → and, else or; customFilter@operator defaults to equal (an unknown one reads equal); two conditions at most, none is no criterion", () => {
    const criteria = (body: string): FilterCriteria | undefined =>
      read(`<sheetData/><autoFilter ref="A1:A9"><filterColumn colId="0">${body}</filterColumn></autoFilter>`).autoFilter!.columns[0]!.criteria;
    expect(criteria(`<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="100"/><customFilter operator="lessThanOrEqual" val="900"/></customFilters>`)).toEqual({
      kind: "custom",
      join: "and",
      conditions: [
        { operator: "greaterThanOrEqual", value: "100" },
        { operator: "lessThanOrEqual", value: "900" },
      ],
    });
    expect(criteria(`<customFilters><customFilter val="a*"/><customFilter operator="notEqual" val=" "/></customFilters>`)).toEqual({
      kind: "custom",
      join: "or",
      conditions: [
        { operator: "equal", value: "a*" },
        { operator: "notEqual", value: " " },
      ],
    });
    expect(criteria(`<customFilters and="0"><customFilter operator="between" val="1"/><customFilter val="2"/><customFilter val="3"/></customFilters>`)).toEqual({
      kind: "custom",
      join: "or",
      conditions: [
        { operator: "equal", value: "1" },
        { operator: "equal", value: "2" },
      ],
    });
    expect(criteria(`<customFilters and="1"/>`)).toBeUndefined();
  });

  it("top10@top (default 1), @percent, @val → top10; @filterVal is ignored", () => {
    const criteria = (body: string): FilterCriteria | undefined =>
      read(`<sheetData/><autoFilter ref="A1:A9"><filterColumn colId="0">${body}</filterColumn></autoFilter>`).autoFilter!.columns[0]!.criteria;
    expect(criteria(`<top10 val="5" filterVal="42"/>`)).toEqual({ kind: "top10", top: true, percent: false, value: 5 });
    expect(criteria(`<top10 top="0" percent="1" val="25" filterVal="7"/>`)).toEqual({ kind: "top10", top: false, percent: true, value: 25 });
  });

  it("dynamicFilter@type → dynamic (val, maxVal and valIso ignored); an unknown type is kept verbatim as preserved, labelled icon", () => {
    const criteria = (body: string): FilterCriteria | undefined =>
      read(`<sheetData/><autoFilter ref="A1:A9"><filterColumn colId="0">${body}</filterColumn></autoFilter>`).autoFilter!.columns[0]!.criteria;
    expect(criteria(`<dynamicFilter type="aboveAverage" val="512.5"/>`)).toEqual({ kind: "dynamic", type: "aboveAverage" });
    expect(criteria(`<dynamicFilter type="thisMonth" val="46296" valIso="2026-10-01T00:00:00" maxVal="46327"/>`)).toEqual({ kind: "dynamic", type: "thisMonth" });
    expect(criteria(`<dynamicFilter type="M12"/>`)).toEqual({ kind: "dynamic", type: "M12" });
    expect(criteria(`<dynamicFilter type="null"/>`)).toEqual({ kind: "preserved", xml: `<dynamicFilter type="null"/>`, label: "icon" });
  });

  it("colorFilter and iconFilter → preserved, the element verbatim, labelled colour and icon", () => {
    const filter = read(
      `<sheetData/><autoFilter ref="A1:B9"><filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>` +
        `<filterColumn colId="1"><iconFilter iconSet="3Arrows" iconId="2"></iconFilter></filterColumn></autoFilter>`,
    ).autoFilter!;
    expect(filter.columns).toEqual([
      { colId: 0, criteria: { kind: "preserved", xml: `<colorFilter dxfId="0"/>`, label: "colour" } },
      { colId: 1, criteria: { kind: "preserved", xml: `<iconFilter iconSet="3Arrows" iconId="2"></iconFilter>`, label: "icon" } },
    ]);
  });

  it("autoFilter/sortState → autoFilter.sort; the worksheet sortState → sortState: ref, caseSensitive, a key per value condition (sortBy cellColor or a customList is none), none for columnSort, sourceXml always", () => {
    const nested = `<sortState ref="A2:D13" caseSensitive="1"><sortCondition descending="1" ref="C2:C13"/><sortCondition ref="A2:A13" sortBy="value"/></sortState>`;
    const content = read(
      `<sheetData/><autoFilter ref="A1:D13">${nested}</autoFilter>` +
        `<sortState ref="F2:H9"><sortCondition ref="G2:G9" sortBy="cellColor" dxfId="0"/><sortCondition ref="H2:H9" customList="Jan,Feb"/><sortCondition descending="1" ref="F2:F9"/></sortState>`,
    );
    expect(content.autoFilter!.sort).toEqual({
      ref: range("A2:D13"),
      keys: [
        { col: 2, descending: true },
        { col: 0, descending: false },
      ],
      caseSensitive: true,
      sourceXml: nested,
    });
    expect(content.sortState).toEqual({
      ref: range("F2:H9"),
      keys: [{ col: 5, descending: true }],
      sourceXml: `<sortState ref="F2:H9"><sortCondition ref="G2:G9" sortBy="cellColor" dxfId="0"/><sortCondition ref="H2:H9" customList="Jan,Feb"/><sortCondition descending="1" ref="F2:F9"/></sortState>`,
    });
    const columns = `<sortState columnSort="1" ref="B2:F3"><sortCondition ref="B2:F2"/></sortState>`;
    expect(read(`<sheetData/>${columns}`).sortState).toEqual({ ref: range("B2:F3"), keys: [], sourceXml: columns });
    expect(readSortStateXml(`<sortState ref="nowhere"/>`)).toBeUndefined();
  });

  it("reads a prefixed part's x:autoFilter and x:sortState like unprefixed ones", () => {
    const root = `<x:worksheet xmlns:x="${MAIN}" xmlns:xr="${XR}">`;
    const filter = `<x:autoFilter ref="A1:B5" xr:uid="${UID}"><x:filterColumn colId="1"><x:customFilters><x:customFilter operator="lessThan" val="3"/></x:customFilters></x:filterColumn></x:autoFilter>`;
    const sort = `<x:sortState ref="A2:B5"><x:sortCondition ref="B2:B5"/></x:sortState>`;
    const content = readWorksheet(`${root}<x:sheetData/>${filter}${sort}</x:worksheet>`, "S", { sharedStrings: [], dateSystem: "1900", budget: new CellBudget(), mdw: 7 })!;
    expect(content.autoFilter).toEqual({
      range: range("A1:B5"),
      columns: [{ colId: 1, criteria: { kind: "custom", join: "or", conditions: [{ operator: "lessThan", value: "3" }] } }],
      sourceXml: filter,
    });
    expect(content.sortState).toEqual({ ref: range("A2:B5"), keys: [{ col: 1, descending: false }], sourceXml: sort });
    // An element of another namespace with a criterion's name is no criterion.
    const foreign = `<autoFilter ref="A1:A3" xmlns:o="urn:other"><filterColumn colId="0"><o:top10 val="1"/></filterColumn></autoFilter>`;
    expect(readAutoFilterXml(foreign)!.columns).toEqual([{ colId: 0 }]);
  });

  it("does not model an autoFilter or sortState whose ref does not parse", () => {
    const content = read(`<sheetData/><autoFilter><filterColumn colId="0"/></autoFilter><sortState ref="?"/>`);
    expect(content.autoFilter).toBeUndefined();
    expect(content.sortState).toBeUndefined();
  });

  it("sheetProtection@sort=\"0\" → allowSort, @autoFilter=\"0\" → allowAutoFilter (absent: not permitted)", () => {
    expect(read(`<sheetData/><sheetProtection sheet="1" sort="0" autoFilter="false"/>`).protection).toEqual({ sheet: true, allowSort: true, allowAutoFilter: true });
    expect(read(`<sheetData/><sheetProtection sheet="1" sort="1"/>`).protection).toEqual({ sheet: true });
  });

  it("readSheetFilterElements: a whole part's root-level autoFilter and sortState, the cells unread", () => {
    const filter = `<autoFilter ref="A1:B3"/>`;
    const sort = `<sortState ref="D2:D3"><sortCondition ref="D2:D3"/></sortState>`;
    const part = worksheet(`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${filter}${sort}`);
    expect(readSheetFilterElements(part)).toEqual({ autoFilter: { range: range("A1:B3"), columns: [], sourceXml: filter }, sortState: { ref: range("D2:D3"), keys: [{ col: 3, descending: false }], sourceXml: sort } });
    expect(readSheetFilterElements(worksheet(`<sheetData/>`))).toEqual({});
  });
});

// --- writing -----------------------------------------------------------------------------------

describe("writing: model → xml → read again", () => {
  const RANGE = range("A1:B7");
  // Numbers only: what the stub string table writes would not read back.
  const data = rowsOf([null, 10], [null, 40], [null, 20], [null, 30], [null, null, { hidden: true }], [null, 50]);
  const cases: [string, AutoFilter, string][] = [
    [
      "values with (Blanks) and date groups",
      {
        range: RANGE,
        columns: [
          {
            colId: 0,
            criteria: {
              kind: "values",
              values: ["East", "A & \"B\" <C>"],
              blanks: true,
              dateGroups: [
                { grouping: "month", year: 2026, month: 3 },
                { grouping: "minute", year: 2025, month: 1, day: 2, hour: 3, minute: 4 },
              ],
            },
          },
        ],
      },
      `<autoFilter ref="A1:B7"><filterColumn colId="0"><filters blank="1"><filter val="East"/><filter val="A &amp; &quot;B&quot; &lt;C&gt;"/>` +
        `<dateGroupItem year="2026" month="3" dateTimeGrouping="month"/><dateGroupItem year="2025" month="1" day="2" hour="3" minute="4" dateTimeGrouping="minute"/></filters></filterColumn></autoFilter>`,
    ],
    [
      "values, nothing checked",
      { range: RANGE, columns: [{ colId: 0, criteria: { kind: "values", values: [], blanks: false, dateGroups: [] } }] },
      `<autoFilter ref="A1:B7"><filterColumn colId="0"><filters/></filterColumn></autoFilter>`,
    ],
    [
      "custom, two conditions joined by and (operator omitted for equal)",
      {
        range: RANGE,
        columns: [
          {
            colId: 1,
            criteria: {
              kind: "custom",
              join: "and",
              conditions: [
                { operator: "equal", value: "1*" },
                { operator: "notEqual", value: " " },
              ],
            },
          },
        ],
      },
      `<autoFilter ref="A1:B7"><filterColumn colId="1"><customFilters and="1"><customFilter val="1*"/><customFilter operator="notEqual" val=" "/></customFilters></filterColumn></autoFilter>`,
    ],
    [
      "custom, one condition (or)",
      { range: RANGE, columns: [{ colId: 1, criteria: { kind: "custom", join: "or", conditions: [{ operator: "greaterThan", value: "25" }] } }] },
      `<autoFilter ref="A1:B7"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="25"/></customFilters></filterColumn></autoFilter>`,
    ],
    [
      "top10 items",
      { range: RANGE, columns: [{ colId: 1, criteria: { kind: "top10", top: true, percent: false, value: 2 } }] },
      `<autoFilter ref="A1:B7"><filterColumn colId="1"><top10 val="2" filterVal="40"/></filterColumn></autoFilter>`,
    ],
    [
      "top10 bottom percent",
      { range: RANGE, columns: [{ colId: 1, criteria: { kind: "top10", top: false, percent: true, value: 50 } }] },
      `<autoFilter ref="A1:B7"><filterColumn colId="1"><top10 top="0" percent="1" val="50" filterVal="20"/></filterColumn></autoFilter>`,
    ],
    [
      "dynamic above average",
      { range: RANGE, columns: [{ colId: 1, criteria: { kind: "dynamic", type: "aboveAverage" } }] },
      `<autoFilter ref="A1:B7"><filterColumn colId="1"><dynamicFilter type="aboveAverage" val="30"/></filterColumn></autoFilter>`,
    ],
    [
      "dynamic date",
      { range: RANGE, columns: [{ colId: 0, criteria: { kind: "dynamic", type: "lastQuarter" } }] },
      `<autoFilter ref="A1:B7"><filterColumn colId="0"><dynamicFilter type="lastQuarter"/></filterColumn></autoFilter>`,
    ],
    [
      "preserved colour and icon filters, verbatim; a hidden button",
      {
        range: RANGE,
        columns: [
          { colId: 0, criteria: { kind: "preserved", xml: `<colorFilter dxfId="3" cellColor="0"/>`, label: "colour" }, hideButton: true },
          { colId: 1, criteria: { kind: "preserved", xml: `<iconFilter iconSet="3Flags"/>`, label: "icon" } },
        ],
      },
      `<autoFilter ref="A1:B7"><filterColumn colId="0" hiddenButton="1"><colorFilter dxfId="3" cellColor="0"/></filterColumn>` +
        `<filterColumn colId="1"><iconFilter iconSet="3Flags"/></filterColumn></autoFilter>`,
    ],
    [
      "a column without criteria, and the sort record",
      {
        range: RANGE,
        columns: [{ colId: 1 }],
        sort: {
          ref: range("A2:B7"),
          keys: [
            { col: 1, descending: true },
            { col: 0, descending: false },
          ],
          caseSensitive: true,
        },
      },
      `<autoFilter ref="A1:B7"><filterColumn colId="1"/><sortState ref="A2:B7" caseSensitive="1"><sortCondition descending="1" ref="B2:B7"/><sortCondition ref="A2:A7"/></sortState></autoFilter>`,
    ],
    ["no columns", { range: range("C3:E9"), columns: [] }, `<autoFilter ref="C3:E9"/>`],
  ];

  for (const [label, filter, xml] of cases) {
    it(`${label}: writes ${xml.length > 70 ? `${xml.slice(0, 70)}…` : xml}, which reads back as the model`, () => {
      const out = written(sheetOf({ rows: data, autoFilter: filter }));
      expect(elementIn(out, "autoFilter")).toBe(xml);
      const body = out.slice(out.indexOf(ROOT) + ROOT.length, -"</worksheet>".length);
      const reread = read(body);
      expect(withoutSource(reread.autoFilter)).toEqual(filter);
      expect(reread.autoFilter!.sourceXml).toBe(xml);
    });
  }

  it("the worksheet sortState: each key's ref its column over the state's rows; written in CT_Worksheet order after the autoFilter, before mergeCells", () => {
    const sortState: SortState = { ref: range("C5:F20"), keys: [{ col: 4, descending: false }, { col: 3, descending: true }] };
    const out = written(sheetOf({ sortState, autoFilter: { range: range("A1:A2"), columns: [] }, merges: [range("H1:I1")] }), `<sheetData/><pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/>`);
    expect(out).toContain(
      `<sheetData/><autoFilter ref="A1:A2"/><sortState ref="C5:F20"><sortCondition ref="E5:E20"/><sortCondition descending="1" ref="D5:D20"/></sortState>` +
        `<mergeCells count="1"><mergeCell ref="H1:I1"/></mergeCells><pageMargins `,
    );
    const reread = read(out.slice(out.indexOf(ROOT) + ROOT.length, -"</worksheet>".length));
    expect(withoutSource(reread.sortState)).toEqual(sortState);
    expect(written(sheetOf({ sortState: { ref: range("A2:A9"), keys: [] } }))).toContain(`<sortState ref="A2:A9"/>`);
  });

  it("top10's filterVal: the value-th largest or smallest numeric data cell, rank max(1, floor(count × value / 100)) for percent; none without a number", () => {
    expect(top10Threshold([10, 40, 20, 30, 50], true, false, 2)).toBe(40);
    expect(top10Threshold([10, 40, 20, 30, 50], false, false, 1)).toBe(10);
    expect(top10Threshold([10, 40, 20, 30, 50], true, false, 9)).toBe(10);
    expect(top10Threshold([10, 40, 20, 30, 50], true, true, 10)).toBe(50);
    expect(top10Threshold([10, 40, 20, 30, 50], true, true, 60)).toBe(30);
    expect(top10Threshold([10, 40, 20, 30, 50], false, true, 40)).toBe(20);
    expect(top10Threshold([], true, false, 10)).toBeUndefined();
    // The header row and the cells outside the range's rows are no data.
    const rows: Row[] = [{ index: 0, cells: [{ col: 1, value: 1000 }] }, ...rowsOf(["a", 5], ["b", "x" as unknown as number]), { index: 9, cells: [{ col: 1, value: 999 }] }];
    const top: AutoFilter = { range: range("A1:B3"), columns: [{ colId: 1, criteria: { kind: "top10", top: true, percent: false, value: 1 } }] };
    expect(elementIn(written(sheetOf({ rows, autoFilter: top })), "top10")).toBe(`<top10 val="1" filterVal="5"/>`);
    const none: AutoFilter = { range: range("A1:A3"), columns: [{ colId: 0, criteria: { kind: "top10", top: true, percent: false, value: 3 } }] };
    expect(elementIn(written(sheetOf({ rows, autoFilter: none })), "top10")).toBe(`<top10 val="3"/>`);
    const average: AutoFilter = { range: range("A1:A3"), columns: [{ colId: 0, criteria: { kind: "dynamic", type: "belowAverage" } }] };
    expect(elementIn(written(sheetOf({ rows, autoFilter: average })), "dynamicFilter")).toBe(`<dynamicFilter type="belowAverage"/>`);
  });

  it("writes a prefixed part's autoFilter and sortState with its prefix", () => {
    const root = `<x:worksheet xmlns:x="${MAIN}">`;
    const filter: AutoFilter = { range: range("A1:B3"), columns: [{ colId: 0, criteria: { kind: "values", values: ["v"], blanks: false, dateGroups: [] } }], sort: { ref: range("A2:B3"), keys: [{ col: 0, descending: false }] } };
    const out = serializeWorksheet(sheetOf({ autoFilter: filter, sortState: { ref: range("D2:D3"), keys: [] } }), `${root}<x:sheetData/></x:worksheet>`, stubContext());
    expect(out).toContain(
      `<x:sheetData/><x:autoFilter ref="A1:B3"><x:filterColumn colId="0"><x:filters><x:filter val="v"/></x:filters></x:filterColumn>` +
        `<x:sortState ref="A2:B3"><x:sortCondition ref="A2:A3"/></x:sortState></x:autoFilter><x:sortState ref="D2:D3"/></x:worksheet>`,
    );
  });
});

// --- the rebuilt sheet -------------------------------------------------------------------------

describe("a rebuilt sheet: sourceXml verbatim, or serialised when it does not fit", () => {
  const FILTER =
    `<autoFilter ref="A1:B4" xr:uid="${UID}"><filterColumn colId="0"><filters blank="1" calendarType="gregorian"><filter val="a"/></filters></filterColumn>` +
    `<filterColumn colId="1"><colorFilter dxfId="0"/></filterColumn><sortState ref="A2:B4"><sortCondition sortBy="cellColor" ref="B2:B4" dxfId="0"/></sortState></autoFilter>`;
  const SORT = `<sortState ref="D2:E4" sortMethod="stroke"><sortCondition ref="E2:E4"/></sortState>`;
  const BODY = `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${FILTER}${SORT}<pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/>`;

  it("re-emits the read autoFilter and sortState byte for byte inside the rebuilt part — xr:uid, calendarType, sortMethod and the cell-colour condition kept", () => {
    const content = read(BODY);
    const out = written(sheetOf({ ...content, rows: [{ index: 0, cells: [{ col: 0, value: 2 }] }] }), BODY);
    expect(out).toContain(`<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>${FILTER}${SORT}<pageMargins `);
  });

  it("serialises the model when the sourceXml does not fit the part: a prefix its root does not bind, another element, an ODS database range", () => {
    const content = read(BODY);
    const plainRoot = `<worksheet xmlns="${MAIN}">`; // no xr: the source's xr:uid cannot go in
    const out = serializeWorksheet(sheetOf({ ...content }), worksheet(`<sheetData/>`, plainRoot), stubContext());
    expect(elementIn(out, "autoFilter")).toBe(
      `<autoFilter ref="A1:B4"><filterColumn colId="0"><filters blank="1"><filter val="a"/></filters></filterColumn>` +
        `<filterColumn colId="1"><colorFilter dxfId="0"/></filterColumn><sortState ref="A2:B4"><sortCondition sortBy="cellColor" ref="B2:B4" dxfId="0"/></sortState></autoFilter>`,
    );
    expect(elementIn(out, "sortState")).toBe(`<sortState ref="A2:B4"><sortCondition sortBy="cellColor" ref="B2:B4" dxfId="0"/></sortState>`);
    const ods: AutoFilter = {
      range: range("A1:C9"),
      columns: [],
      sourceXml: `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.C9" table:display-filter-buttons="true"/>`,
    };
    expect(elementIn(written(sheetOf({ autoFilter: ods })), "autoFilter")).toBe(`<autoFilter ref="A1:C9"/>`);
    expect(fitsPart(`<sortState ref="A1"/>`, ROOT, "autoFilter")).toBe(false);
    expect(fitsPart(`<autoFilter ref="A1"/><autoFilter ref="A2"/>`, ROOT, "autoFilter")).toBe(false);
    expect(fitsPart(`<autoFilter ref="A1" xr:uid="x"/>`, ROOT, "autoFilter")).toBe(true);
    expect(fitsPart(`<autoFilter ref="A1" xr:uid="x"/>`, `<worksheet xmlns="${MAIN}">`, "autoFilter")).toBe(false);
    expect(fitsPart(`<autoFilter ref="A1"`, ROOT, "autoFilter")).toBe(false);
  });

  it("removes a source autoFilter and sortState the model no longer has; keeps one the reader could not model", () => {
    expect(written(sheetOf(), BODY)).toContain(`<sheetData/><pageMargins `);
    const odd = `<autoFilter><filterColumn colId="0"/></autoFilter><sortState ref="?"/>`;
    expect(written(sheetOf(), `<sheetData/>${odd}`)).toContain(`<sheetData/>${odd}</worksheet>`);
    // The model's own replaces even an unmodellable source element.
    expect(written(sheetOf({ autoFilter: { range: range("A1:A2"), columns: [] } }), `<sheetData/>${odd}`)).toContain(`<sheetData/><autoFilter ref="A1:A2"/><sortState ref="?"/></worksheet>`);
  });
});

// --- rows and filterMode -------------------------------------------------------------------------

describe("rows and sheetPr@filterMode", () => {
  it("writes hidden=\"1\" for a hidden and for a filter-hidden row — a filter-hidden row with no cells too, which the reader stores", () => {
    const rows: Row[] = [
      { index: 1, cells: [{ col: 0, value: 1 }], hidden: true },
      { index: 2, cells: [{ col: 0, value: 2 }], filterHidden: true },
      { index: 3, cells: [], filterHidden: true },
      { index: 4, cells: [], hidden: true, filterHidden: true },
    ];
    const out = written(sheetOf({ rows }));
    expect(out).toContain(`<row r="2" hidden="1"><c r="A2"><v>1</v></c></row><row r="3" hidden="1"><c r="A3"><v>2</v></c></row><row r="4" hidden="1"/><row r="5" hidden="1"/>`);
    expect(read(`<sheetData><row r="4" hidden="1"/></sheetData>`).rows).toEqual([{ index: 3, cells: [], hidden: true }]);
  });

  const hiddenByFilter: Row[] = [{ index: 2, cells: [], filterHidden: true }];

  it("sets filterMode=\"1\" exactly when some row is filter-hidden: spliced into the source sheetPr, its children untouched", () => {
    const sheetPr = `<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>`;
    const out = written(sheetOf({ rows: hiddenByFilter }), `${sheetPr}<dimension ref="A1"/><sheetData/>`);
    expect(out).toContain(`${ROOT}<sheetPr codeName="Sheet1" filterMode="1"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr><dimension `);
    // A user-hidden row alone is no filter.
    expect(written(sheetOf({ rows: [{ index: 2, cells: [], hidden: true }] }), `${sheetPr}<sheetData/>`)).toContain(`${ROOT}${sheetPr}`);
  });

  it("removes filterMode when no row is filter-hidden; leaves one already saying so as written", () => {
    expect(written(sheetOf(), `<sheetPr filterMode="1" codeName="S"/><sheetData/>`)).toContain(`${ROOT}<sheetPr codeName="S"/>`);
    expect(written(sheetOf(), `<sheetPr filterMode="1"/><sheetData/>`)).toContain(`${ROOT}<sheetPr/>`);
    expect(written(sheetOf({ rows: hiddenByFilter }), `<sheetPr filterMode="true"/><sheetData/>`)).toContain(`${ROOT}<sheetPr filterMode="true"/>`);
  });

  it("splices a self-closing sheetPr, and makes <sheetPr filterMode=\"1\"/> the first root child when the part has none", () => {
    expect(written(sheetOf({ rows: hiddenByFilter }), `<sheetPr codeName="S"/><sheetData/>`)).toContain(`${ROOT}<sheetPr codeName="S" filterMode="1"/>`);
    expect(written(sheetOf({ rows: hiddenByFilter }), `<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>`)).toContain(
      `${ROOT}<sheetPr filterMode="1"/><dimension ref="A1"/><sheetViews>`,
    );
    // A new sheet: the empty-sheet template, with and without a tab colour.
    expect(serializeWorksheet(sheetOf({ rows: hiddenByFilter }), undefined, stubContext())).toContain(`<sheetPr filterMode="1"/><dimension ref="A1"/>`);
    expect(serializeWorksheet(sheetOf({ rows: hiddenByFilter, tabColor: { rgb: "#ff0000" } }), undefined, stubContext())).toMatch(/<sheetPr filterMode="1"><tabColor [^>]*\/><\/sheetPr><dimension /);
  });
});

// --- _xlnm._FilterDatabase ---------------------------------------------------------------------

interface BookSheet {
  name: string;
  /** The worksheet root's children. */
  body?: string;
}

/** A minimal Excel-shaped package: the sheets, inline strings, `workbookTail`
 * between `</sheets>` and `<calcPr>`. */
async function book(sheets: readonly BookSheet[], workbookTail = ""): Promise<Uint8Array> {
  const files: Record<string, string> = {
    "[Content_Types].xml":
      `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
      `</Types>`,
    "_rels/.rels": `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml":
      `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>` +
      sheets.map((sheet, i) => `<sheet name="${sheet.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
      `</sheets>${workbookTail}<calcPr calcId="191029"/></workbook>`,
    "xl/_rels/workbook.xml.rels":
      `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
      `</Relationships>`,
  };
  sheets.forEach((sheet, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = worksheet(sheet.body ?? `<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/>`);
  });
  return zip(Object.entries(files).map(([name, text]) => ({ name, data: encoder.encode(text) })));
}

const cells = `<dimension ref="A1:B3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
  `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>` +
  `<row r="2"><c r="A2" t="inlineStr"><is><t>East</t></is></c><c r="B2"><v>5</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>West</t></is></c><c r="B3"><v>7</v></c></row></sheetData>`;
const FILTERED = `${cells}<autoFilter ref="A1:B3" xr:uid="${UID}"/>`;
const NAME_0 = `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sales!$A$1:$B$3</definedName>`;
const PRINT_0 = `<definedName name="_xlnm.Print_Area" localSheetId="0">Sales!$A$1:$B$3</definedName>`;

async function open(bytes: Uint8Array): Promise<SpreadsheetModel> {
  return readSpreadsheetBytes(bytes, "xlsx", "t");
}

async function workbookXml(model: SpreadsheetModel): Promise<string> {
  const parts = await unzip(await writeXlsx(model));
  return decoder.decode(parts.get("xl/workbook.xml"));
}

const definedNamesOf = (xml: string): string => /<definedNames>.*<\/definedNames>|<definedNames\/>/.exec(xml)?.[0] ?? "";

describe("_xlnm._FilterDatabase in xl/workbook.xml", () => {
  it("an untouched filtered workbook downloads byte-identical, part for part", async () => {
    const bytes = await book([{ name: "Sales", body: FILTERED }, { name: "Other" }], `<definedNames>${NAME_0}${PRINT_0}</definedNames>`);
    const source = await unzip(bytes);
    const out = await unzip(await writeXlsx(await open(bytes)));
    expect([...out.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("a dirty sheet whose autofilter did not move leaves the workbook part as the source wrote it", async () => {
    const bytes = await book([{ name: "Sales", body: FILTERED }, { name: "Other" }], `<definedNames>${NAME_0}${PRINT_0}</definedNames>`);
    const model = await open(bytes);
    model.sheets[0] = { ...model.sheets[0]!, rows: model.sheets[0]!.rows.slice(0, 2), dirty: true };
    expect(await workbookXml(model)).toBe(decoder.decode((await unzip(bytes)).get("xl/workbook.xml")));
  });

  it("moved: a dirty sheet's autofilter at a new range rewrites the name's text, its tag kept; the other names untouched", async () => {
    const model = await open(await book([{ name: "Sales", body: FILTERED }, { name: "Other" }], `<definedNames>${NAME_0}${PRINT_0}</definedNames>`));
    model.sheets[0] = { ...model.sheets[0]!, autoFilter: { range: range("A3:B5"), columns: [] }, dirty: true };
    expect(definedNamesOf(await workbookXml(model))).toBe(
      `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sales!$A$3:$B$5</definedName>${PRINT_0}</definedNames>`,
    );
  });

  it("removed: a dirty sheet without an autofilter loses its name; a definedNames left empty stays", async () => {
    const model = await open(await book([{ name: "Sales", body: FILTERED }], `<definedNames>${NAME_0}</definedNames>`));
    model.sheets[0] = { ...model.sheets[0]!, autoFilter: undefined, dirty: true };
    const xml = await workbookXml(model);
    expect(xml).not.toContain("_FilterDatabase");
    expect(definedNamesOf(xml)).toBe(`<definedNames></definedNames>`);
  });

  it("added: a sheet gaining an autofilter gains a hidden name, first in definedNames, its sheet name quoted as a reference quotes it", async () => {
    const model = await open(await book([{ name: "Sales", body: FILTERED }, { name: "My Data", body: cells }], `<definedNames>${NAME_0}${PRINT_0}</definedNames>`));
    model.sheets[1] = { ...model.sheets[1]!, autoFilter: { range: range("A1:B3"), columns: [] }, dirty: true };
    expect(definedNamesOf(await workbookXml(model))).toBe(
      `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'My Data'!$A$1:$B$3</definedName>${NAME_0}${PRINT_0}</definedNames>`,
    );
  });

  it("creates <definedNames> after <sheets> when the part has none, and opens up a self-closing one", async () => {
    const model = await open(await book([{ name: "Sales", body: cells }]));
    model.sheets[0] = { ...model.sheets[0]!, autoFilter: { range: range("A1:B3"), columns: [] }, dirty: true };
    expect(await workbookXml(model)).toContain(`</sheets><definedNames>${NAME_0}</definedNames><calcPr calcId="191029"/>`);
    const closed = await open(await book([{ name: "Sales", body: cells }], `<definedNames/>`));
    closed.sheets[0] = { ...closed.sheets[0]!, autoFilter: { range: range("A1:B3"), columns: [] }, dirty: true };
    expect(await workbookXml(closed)).toContain(`</sheets><definedNames>${NAME_0}</definedNames><calcPr`);
  });

  it("an existing name naming the range another way (unquoted, no $) is left as written", async () => {
    expect(filterDatabaseNames("'Sales'!A1:B3", "sales", range("A1:B3"))).toBe(true);
    expect(filterDatabaseNames("Sales!$A$1:$B$4", "Sales", range("A1:B3"))).toBe(false);
    expect(filterDatabaseNames("Other!$A$1:$B$3", "Sales", range("A1:B3"))).toBe(false);
    expect(filterDatabaseNames("#REF!", "Sales", range("A1:B3"))).toBe(false);
    expect(filterDatabaseText("It's", range("C2:C2"))).toBe(`'It''s'!$C$2:$C$2`);
    const loose = `<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Sales'!A1:B3</definedName>`;
    const model = await open(await book([{ name: "Sales", body: FILTERED }], `<definedNames>${loose}</definedNames>`));
    model.sheets[0] = { ...model.sheets[0]!, dirty: true };
    expect(definedNamesOf(await workbookXml(model))).toBe(`<definedNames>${loose}</definedNames>`);
  });

  it("moving a sheet renumbers its name with SPEC-121's localSheetId rule — once, never duplicated", async () => {
    const model = await open(await book([{ name: "Sales", body: FILTERED }, { name: "Other" }, { name: "Third" }], `<definedNames>${NAME_0}${PRINT_0}</definedNames>`));
    moveSheet(model, 0, 2);
    expect(definedNamesOf(await workbookXml(model))).toBe(
      `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">Sales!$A$1:$B$3</definedName>` +
        `<definedName name="_xlnm.Print_Area" localSheetId="2">Sales!$A$1:$B$3</definedName></definedNames>`,
    );
  });

  it("renaming a filtered sheet rewrites its name's text to the new name", async () => {
    const model = await open(await book([{ name: "Sales", body: FILTERED }], `<definedNames>${NAME_0}</definedNames>`));
    renameSheet(model, 0, "Q4 Sales");
    expect(definedNamesOf(await workbookXml(model))).toBe(`<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Q4 Sales'!$A$1:$B$3</definedName></definedNames>`);
  });

  it("deleting a filtered sheet removes its name and renumbers the later sheet's", async () => {
    const second = `<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Other!$A$1:$B$3</definedName>`;
    const model = await open(await book([{ name: "Sales", body: FILTERED }, { name: "Other", body: FILTERED }], `<definedNames>${NAME_0}${second}</definedNames>`));
    deleteSheet(model, 0);
    expect(definedNamesOf(await workbookXml(model))).toBe(`<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Other!$A$1:$B$3</definedName></definedNames>`);
  });

  it("duplicating a filtered sheet gives the copy its own name from its autofilter, and the origin keeps exactly one", async () => {
    const model = await open(await book([{ name: "Sales", body: FILTERED }, { name: "Other" }], `<definedNames>${NAME_0}${PRINT_0}</definedNames>`));
    duplicateSheet(model, 0);
    const names = definedNamesOf(await workbookXml(model));
    expect(names).toBe(
      `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'Sales (2)'!$A$1:$B$3</definedName>${NAME_0}${PRINT_0}` +
        `<definedName name="_xlnm.Print_Area" localSheetId="1">'Sales (2)'!$A$1:$B$3</definedName></definedNames>`,
    );
    expect(names.match(/_FilterDatabase/g)).toHaveLength(2);
  });

  it("a from-scratch workbook writes one hidden name per filtered sheet", async () => {
    const model = createBlankWorkbook();
    model.sheets[0] = { ...model.sheets[0]!, autoFilter: { range: range("B2:C9"), columns: [] } };
    const xml = await workbookXml(model);
    expect(xml).toContain(`</sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$B$2:$C$9</definedName></definedNames><calcPr`);
    expect(await workbookXml(createBlankWorkbook())).not.toContain("definedNames");
  });
});

// --- the read classification ---------------------------------------------------------------------

/** Region in A, Amount in B; rows 2–7 data, row 9 outside the range. */
const regions = (filterColumn: string, rows: Array<[string, number, boolean]>): string =>
  `<dimension ref="A1:B9"/><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c></row>` +
  rows.map(([text, amount, hidden], i) => `<row r="${i + 2}"${hidden ? ` hidden="1"` : ""}><c r="A${i + 2}" t="inlineStr"><is><t>${text}</t></is></c><c r="B${i + 2}"><v>${amount}</v></c></row>`).join("") +
  `<row r="9" hidden="1"><c r="A9" t="inlineStr"><is><t>South</t></is></c></row></sheetData><autoFilter ref="A1:B7">${filterColumn}</autoFilter>`;
const flags = (sheet: Sheet): Record<number, string> =>
  Object.fromEntries(sheet.rows.filter((row) => row.hidden || row.filterHidden).map((row) => [row.index + 1, row.filterHidden ? (row.hidden ? "both" : "filter") : "user"]));

describe("reading classifies the hidden rows inside an autofilter (filter.ts's read rule)", () => {
  it("a hidden row the criteria reject → filterHidden; one they pass → hidden; a shown row that fails stays shown; a hidden row outside the range → hidden", async () => {
    const body = regions(`<filterColumn colId="0"><filters><filter val="East"/><filter val="West"/></filters></filterColumn>`, [
      ["East", 1, false],
      ["South", 2, true],
      ["West", 3, true],
      ["South", 4, false],
      ["North", 5, true],
      ["East", 6, false],
    ]);
    const sheet = (await open(await book([{ name: "Sales", body }]))).sheets[0]!;
    expect(flags(sheet)).toEqual({ 3: "filter", 4: "user", 6: "filter", 9: "user" });
  });

  it("with a preserved criterion every hidden row inside the range is filterHidden", async () => {
    const body = regions(`<filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>`, [
      ["East", 1, true],
      ["West", 2, false],
      ["East", 3, true],
    ]);
    const sheet = (await open(await book([{ name: "Colours", body }]))).sheets[0]!;
    expect(flags(sheet)).toEqual({ 2: "filter", 4: "filter", 9: "user" });
  });

  it("without an autofilter every hidden row stays user-hidden", async () => {
    const body = regions("", [["East", 1, true]]).replace(/<autoFilter[^]*<\/autoFilter>/, "");
    const sheet = (await open(await book([{ name: "S", body }]))).sheets[0]!;
    expect(flags(sheet)).toEqual({ 2: "user", 9: "user" });
  });

  it("filter → save → reopen: the autoFilter, the sortState and every row's hidden/filterHidden come back (sourceXml aside)", async () => {
    const model = await open(await book([{ name: "Sales", body: regions("", [["East", 1, false], ["West", 2, false], ["East", 3, false]]).replace(/<row r="9"[^]*?<\/row>/, "") }]));
    const autoFilter: AutoFilter = {
      range: range("A1:B4"),
      columns: [{ colId: 0, criteria: { kind: "values", values: ["East"], blanks: false, dateGroups: [] } }],
      sort: { ref: range("A2:B4"), keys: [{ col: 1, descending: true }] },
    };
    const sortState: SortState = { ref: range("D2:D9"), keys: [{ col: 3, descending: false }], caseSensitive: true };
    const sheet = model.sheets[0]!;
    model.sheets[0] = {
      ...sheet,
      autoFilter,
      sortState,
      rows: sheet.rows.map((row) => (row.index === 2 ? { ...row, filterHidden: true } : row.index === 3 ? { ...row, hidden: true } : row)),
      dirty: true,
    };
    model.structureDirty = true;
    const reread = (await open(await writeXlsx(model))).sheets[0]!;
    expect(withoutSource(reread.autoFilter)).toEqual(autoFilter);
    expect(withoutSource(reread.sortState)).toEqual(sortState);
    expect(flags(reread)).toEqual({ 3: "filter", 4: "user" });
  });
});

// --- spec128-filter.xlsx -------------------------------------------------------------------------

describe("spec128-filter.xlsx", () => {
  const F = SPEC128_FILTER_XLSX;

  it("reads Sales' values and custom filters with its sort record, Colours' colour filter, Log's cell-colour sortState (no key)", async () => {
    const model = await open(await buildSpec128FilterXlsxBytes());
    const [sales, colours, log] = model.sheets;
    expect(sales!.autoFilter).toEqual({
      range: F.sales.autoFilter.range,
      columns: [
        { colId: F.sales.autoFilter.region.colId, criteria: { kind: "values", values: [...F.sales.autoFilter.region.values], blanks: true, dateGroups: [] } },
        { colId: F.sales.autoFilter.amount.colId, criteria: { kind: "custom", join: "and", conditions: F.sales.autoFilter.amount.conditions.map((condition) => ({ ...condition })) } },
      ],
      sort: { ref: F.sales.autoFilter.sort.range, keys: F.sales.autoFilter.sort.keys.map((key) => ({ ...key })), sourceXml: expect.any(String) },
      sourceXml: F.sales.autoFilter.xml,
    });
    expect(colours!.autoFilter!.columns).toEqual([{ colId: 0, criteria: { kind: "preserved", xml: F.colours.autoFilter.colorFilter.xml, label: "colour" } }]);
    expect(log!.sortState).toEqual({ ref: F.log.sortState.range, keys: [], sourceXml: F.log.sortState.xml });
    expect(sales!.rows.filter((row) => row.filterHidden).map((row) => row.index)).toEqual([...F.sales.filterHiddenIndices]);
    expect(sales!.rows.filter((row) => row.hidden).map((row) => row.index)).toEqual([F.sales.userHiddenIndex]);
    expect(colours!.rows.filter((row) => row.filterHidden).map((row) => row.index)).toEqual([...F.colours.filterHiddenIndices]);
  });

  it("an edited Sales sheet re-emits its autoFilter and Log's sortState verbatim, keeps filterMode and both _FilterDatabase names", async () => {
    const bytes = await buildSpec128FilterXlsxBytes();
    const source = await unzip(bytes);
    const model = await open(bytes);
    for (const index of [0, 2]) {
      const sheet = model.sheets[index]!;
      model.sheets[index] = { ...sheet, rows: sheet.rows.map((row) => (row.index === 1 ? { ...row, heightPt: 30 } : row)), dirty: true };
    }
    const out = await unzip(await writeXlsx(model));
    const sales = decoder.decode(out.get(F.sales.partName));
    expect(sales).toContain(F.sales.autoFilter.xml);
    expect(sales).toMatch(/<sheetPr[^>]*filterMode="1"/);
    expect(decoder.decode(out.get(F.log.partName))).toContain(F.log.sortState.xml);
    // Both names already name their autofilters: left as Excel wrote them.
    expect(definedNamesOf(decoder.decode(out.get("xl/workbook.xml")))).toBe(definedNamesOf(decoder.decode(source.get("xl/workbook.xml"))));
    expect(definedNamesOf(decoder.decode(out.get("xl/workbook.xml"))).match(/_FilterDatabase/g)).toHaveLength(2);
    const reread = await open(await writeXlsx(model));
    expect(reread.sheets.map((sheet) => [sheet.autoFilter, sheet.sortState])).toEqual(model.sheets.map((sheet) => [sheet.autoFilter, sheet.sortState]));
  });
});
