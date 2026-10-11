// SPEC-128 "ODS wire mapping" (ods/databaseRanges.ts): every row of the
// operator table both ways — read, written as exact XML, and read back to the
// same criteria — the rows that only read (or-of-equals, empty, a text `=`
// "") and the ones with no ODS form; the unmapped trees kept verbatim; the
// sort record; which database range is a sheet's autofilter (the claim);
// the reader's wiring and the rows read exactly as `table:visibility` says;
// the writer — an untouched workbook byte-identical, an unchanged
// autofilter's element verbatim inside a rebuilt content.xml, regeneration
// after a criteria change (in place, other database ranges and the claimed
// element's unmodelled attributes and children kept), a new filter
// (appended, or a new container placed after the named expressions), a
// removal (an emptied container removed), a structural edit moving it once
// (through the model), a verbatim filter's criteria range moved once (the
// model's copy, by a structural edit and a move), a rename, a deletion renumbering an anonymous range,
// a duplicate's own range, from scratch, write → read;
// and the table writer's visibility and its repeat-compression key.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import type { AutoFilter, FilterColumn, FilterCriteria } from "../filter";
import type { Row, Sheet, SpreadsheetModel } from "../model";
import { SPREADSHEET_LOCALES } from "../numberFormat/locale";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook, deleteSheet, duplicateSheet, renameSheet } from "../sheets";
import { moveRange } from "../move";
import { applyStructuralEdit } from "../structure";
import {
  odsAutoFilterClaims,
  odsDatabaseRangeXml,
  odsDatabaseRanges,
  odsFilterXml,
  odsOwnedDatabaseRangeSpans,
  odsSortXml,
  odsTargetRangeAddress,
  readOdsAutoFilters,
} from "./databaseRanges";
import { parseOds } from "./reader";
import { OdsStyleRegistry } from "./styles";
import { serializeTable } from "./tableXml";
import { writeOds } from "./writer";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

// --- the pure mapping ------------------------------------------------------------------------

const RANGE = { top: 0, left: 0, bottom: 4, right: 3 };
const ANONYMOUS_HEAD = `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.D5" table:display-filter-buttons="true">`;
const databaseRanges = (...elements: string[]): string => `<table:database-ranges>${elements.join("")}</table:database-ranges>`;
const anonymous = (body: string): string => `${ANONYMOUS_HEAD}${body}</table:database-range>`;
const condition = (field: number, attrs: string): string => `<table:filter-condition table:field-number="${field}" ${attrs}/>`;

/** The autofilter Sheet1's claimed element reads as. */
function readOne(element: string, tables: readonly string[] = ["Sheet1"]): AutoFilter | undefined {
  return readOdsAutoFilters(databaseRanges(element), tables).get(0);
}
/** The columns a `table:filter` reads as (undefined: kept verbatim). */
function columnsOf(filter: string): FilterColumn[] | undefined {
  const autoFilter = readOne(anonymous(filter))!;
  return autoFilter.odsUnmappedFilter === undefined ? autoFilter.columns : undefined;
}
const filterOf = (columns: FilterColumn[]): string => odsFilterXml({ range: RANGE, columns });
const custom = (join: "and" | "or", ...conditions: [string, string][]): FilterCriteria =>
  ({ kind: "custom", join, conditions: conditions.map(([operator, value]) => ({ operator, value })) }) as FilterCriteria;
const values = (list: string[], blanks = false): FilterCriteria => ({ kind: "values", values: list, blanks, dateGroups: [] });

/** Rows of the operator table that map both ways: the wire `table:filter`
 * and the columns, written exactly so and read back the same. */
const BOTH_WAYS: [string, string, FilterColumn[]][] = [
  [
    "= with filter-set-items, an empty item for (Blanks)",
    `<table:filter><table:filter-condition table:field-number="0" table:value="East" table:operator="="><table:filter-set-item table:value="East"/><table:filter-set-item table:value="North"/><table:filter-set-item table:value=""/></table:filter-condition></table:filter>`,
    [{ colId: 0, criteria: values(["East", "North"], true) }],
  ],
  [
    "(Blanks) alone",
    `<table:filter><table:filter-condition table:field-number="2" table:value="" table:operator="="><table:filter-set-item table:value=""/></table:filter-condition></table:filter>`,
    [{ colId: 2, criteria: values([], true) }],
  ],
  [
    "= on a number",
    `<table:filter>${condition(1, `table:value="100" table:operator="=" table:data-type="number"`)}</table:filter>`,
    [{ colId: 1, criteria: custom("and", ["equal", "100"]) }],
  ],
  ["!= on a number", `<table:filter>${condition(1, `table:value="-2.5" table:operator="!=" table:data-type="number"`)}</table:filter>`, [{ colId: 1, criteria: custom("and", ["notEqual", "-2.5"]) }]],
  ["< on a number", `<table:filter>${condition(1, `table:value="1E3" table:operator="&lt;" table:data-type="number"`)}</table:filter>`, [{ colId: 1, criteria: custom("and", ["lessThan", "1E3"]) }]],
  ["> on a number", `<table:filter>${condition(1, `table:value="7" table:operator="&gt;" table:data-type="number"`)}</table:filter>`, [{ colId: 1, criteria: custom("and", ["greaterThan", "7"]) }]],
  ["<= on a number", `<table:filter>${condition(1, `table:value="7" table:operator="&lt;=" table:data-type="number"`)}</table:filter>`, [{ colId: 1, criteria: custom("and", ["lessThanOrEqual", "7"]) }]],
  [">= on a number", `<table:filter>${condition(1, `table:value=".5" table:operator="&gt;=" table:data-type="number"`)}</table:filter>`, [{ colId: 1, criteria: custom("and", ["greaterThanOrEqual", ".5"]) }]],
  ["= on text", `<table:filter>${condition(0, `table:value="East" table:operator="="`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["equal", "East"]) }]],
  ["< on text", `<table:filter>${condition(0, `table:value="M" table:operator="&lt;"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["lessThan", "M"]) }]],
  [
    "= on text with a wildcard that is no begins/ends/contains pattern, verbatim",
    `<table:filter>${condition(0, `table:value="E?st" table:operator="="`)}</table:filter>`,
    [{ colId: 0, criteria: custom("and", ["equal", "E?st"]) }],
  ],
  [
    "two conditions on one field inside the filter-and: one custom, and",
    `<table:filter><table:filter-and>${condition(1, `table:value="100" table:operator="&gt;=" table:data-type="number"`)}${condition(1, `table:value="500" table:operator="&lt;=" table:data-type="number"`)}</table:filter-and></table:filter>`,
    [{ colId: 1, criteria: custom("and", ["greaterThanOrEqual", "100"], ["lessThanOrEqual", "500"]) }],
  ],
  [
    "a filter-or of two conditions on one field: one custom, or",
    `<table:filter><table:filter-or>${condition(1, `table:value="10" table:operator="&lt;" table:data-type="number"`)}${condition(1, `table:value="90" table:operator="&gt;" table:data-type="number"`)}</table:filter-or></table:filter>`,
    [{ colId: 1, criteria: custom("or", ["lessThan", "10"], ["greaterThan", "90"]) }],
  ],
  ["begins-with", `<table:filter>${condition(0, `table:value="Ea" table:operator="begins-with"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["equal", "Ea*"]) }]],
  ["ends-with", `<table:filter>${condition(0, `table:value="st" table:operator="ends-with"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["equal", "*st"]) }]],
  ["contains, its wildcards escaped", `<table:filter>${condition(0, `table:value="a*b?c~" table:operator="contains"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["equal", "*a~*b~?c~~*"]) }]],
  ["does-not-begin-with", `<table:filter>${condition(0, `table:value="Ea" table:operator="does-not-begin-with"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["notEqual", "Ea*"]) }]],
  ["does-not-end-with", `<table:filter>${condition(0, `table:value="st" table:operator="does-not-end-with"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["notEqual", "*st"]) }]],
  ["does-not-contain", `<table:filter>${condition(0, `table:value="as" table:operator="does-not-contain"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["notEqual", "*as*"]) }]],
  ["!empty: Excel's non-blank encoding", `<table:filter>${condition(0, `table:value="" table:operator="!empty"`)}</table:filter>`, [{ colId: 0, criteria: custom("and", ["notEqual", " "]) }]],
  [
    "top values",
    `<table:filter>${condition(1, `table:value="3" table:operator="top values" table:data-type="number"`)}</table:filter>`,
    [{ colId: 1, criteria: { kind: "top10", top: true, percent: false, value: 3 } }],
  ],
  [
    "bottom values",
    `<table:filter>${condition(1, `table:value="2" table:operator="bottom values" table:data-type="number"`)}</table:filter>`,
    [{ colId: 1, criteria: { kind: "top10", top: false, percent: false, value: 2 } }],
  ],
  [
    "top percent",
    `<table:filter>${condition(1, `table:value="25" table:operator="top percent" table:data-type="number"`)}</table:filter>`,
    [{ colId: 1, criteria: { kind: "top10", top: true, percent: true, value: 25 } }],
  ],
  [
    "bottom percent",
    `<table:filter>${condition(1, `table:value="10" table:operator="bottom percent" table:data-type="number"`)}</table:filter>`,
    [{ colId: 1, criteria: { kind: "top10", top: false, percent: true, value: 10 } }],
  ],
  [
    "criteria on several columns: a filter-and, ascending colId",
    `<table:filter><table:filter-and><table:filter-condition table:field-number="0" table:value="East" table:operator="="><table:filter-set-item table:value="East"/></table:filter-condition>` +
      `${condition(1, `table:value="10" table:operator="&gt;" table:data-type="number"`)}${condition(3, `table:value="x" table:operator="begins-with"`)}</table:filter-and></table:filter>`,
    [
      { colId: 0, criteria: values(["East"]) },
      { colId: 1, criteria: custom("and", ["greaterThan", "10"]) },
      { colId: 3, criteria: custom("and", ["equal", "x*"]) },
    ],
  ],
  [
    "an or column beside an and column",
    `<table:filter><table:filter-and><table:filter-or>${condition(0, `table:value="A" table:operator="begins-with"`)}${condition(0, `table:value="b" table:operator="ends-with"`)}</table:filter-or>` +
      `${condition(2, `table:value="1" table:operator="&gt;" table:data-type="number"`)}${condition(2, `table:value="9" table:operator="&lt;" table:data-type="number"`)}</table:filter-and></table:filter>`,
    [
      { colId: 0, criteria: custom("or", ["equal", "A*"], ["equal", "*b"]) },
      { colId: 2, criteria: custom("and", ["greaterThan", "1"], ["lessThan", "9"]) },
    ],
  ],
];

describe("the operator table, both ways", () => {
  it.each(BOTH_WAYS)("%s", (_label, wire, columns) => {
    expect(columnsOf(wire)).toEqual(columns);
    expect(filterOf(columns)).toBe(wire);
    expect(columnsOf(filterOf(columns))).toEqual(columns);
  });

  it("escapes values, a multi-line cell's line breaks as character references, so they read back", () => {
    const columns: FilterColumn[] = [{ colId: 0, criteria: values(["a & <b>", "two\nlines\tand \"tab\""]) }];
    const wire = filterOf(columns);
    expect(wire).toContain(`<table:filter-set-item table:value="a &amp; &lt;b&gt;"/><table:filter-set-item table:value="two&#10;lines&#9;and &quot;tab&quot;"/>`);
    expect(columnsOf(wire)).toEqual(columns);
  });

  it("writes a pure pattern as begins/ends/contains even when it came from elsewhere (xlsx's customFilter)", () => {
    expect(filterOf([{ colId: 0, criteria: custom("and", ["equal", "*Ea~?st*"]) }])).toBe(`<table:filter>${condition(0, `table:value="Ea?st" table:operator="contains"`)}</table:filter>`);
    expect(filterOf([{ colId: 0, criteria: custom("and", ["equal", "*"]) }])).toBe(`<table:filter>${condition(0, `table:value="*" table:operator="="`)}</table:filter>`);
    expect(filterOf([{ colId: 0, criteria: custom("and", ["notEqual", "a*b*"]) }])).toBe(`<table:filter>${condition(0, `table:value="a*b*" table:operator="!="`)}</table:filter>`);
  });
});

describe("rows that only read", () => {
  it("reads a filter-or of = on one field (and empty) as values, a text = \"\" and empty as (Blanks)", () => {
    expect(columnsOf(`<table:filter><table:filter-or>${condition(0, `table:value="East" table:operator="="`)}${condition(0, `table:value="West" table:operator="="`)}${condition(0, `table:value="" table:operator="empty"`)}</table:filter-or></table:filter>`)).toEqual([
      { colId: 0, criteria: values(["East", "West"], true) },
    ]);
    expect(columnsOf(`<table:filter>${condition(1, `table:value="" table:operator="empty"`)}</table:filter>`)).toEqual([{ colId: 1, criteria: values([], true) }]);
    expect(columnsOf(`<table:filter>${condition(1, `table:value="" table:operator="="`)}</table:filter>`)).toEqual([{ colId: 1, criteria: values([], true) }]);
    // A field's filter-or nested in the and; set items merge with plain equalities.
    expect(
      columnsOf(
        `<table:filter><table:filter-and><table:filter-or><table:filter-condition table:field-number="0" table:value="A" table:operator="="><table:filter-set-item table:value="A"/></table:filter-condition>` +
          `${condition(0, `table:value="B" table:operator="="`)}</table:filter-or>${condition(1, `table:value="5" table:operator="&gt;" table:data-type="number"`)}</table:filter-and></table:filter>`,
      ),
    ).toEqual([
      { colId: 0, criteria: values(["A", "B"]) },
      { colId: 1, criteria: custom("and", ["greaterThan", "5"]) },
    ]);
  });

  it("known non-identity: a custom or of two plain-text equals reads back as their values", () => {
    const wire = filterOf([{ colId: 0, criteria: custom("or", ["equal", "East"], ["equal", "West"]) }]);
    expect(wire).toBe(`<table:filter><table:filter-or>${condition(0, `table:value="East" table:operator="="`)}${condition(0, `table:value="West" table:operator="="`)}</table:filter-or></table:filter>`);
    expect(columnsOf(wire)).toEqual([{ colId: 0, criteria: values(["East", "West"]) }]);
  });

  it("an empty filter element holds no criteria", () => {
    expect(columnsOf(`<table:filter/>`)).toEqual([]);
    expect(readOne(`<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.D5" table:display-filter-buttons="true"/>`)!.columns).toEqual([]);
  });
});

describe("a tree that does not decompose per column: kept verbatim", () => {
  const UNMAPPED: [string, string][] = [
    ["match", `<table:filter>${condition(0, `table:value="^E" table:operator="match"`)}</table:filter>`],
    ["!match", `<table:filter>${condition(0, `table:value="^E" table:operator="!match"`)}</table:filter>`],
    ["a colour condition", `<table:filter>${condition(0, `table:value="#ff0000" table:operator="=" loext:data-type="background-color"`)}</table:filter>`],
    ["another data type", `<table:filter>${condition(0, `table:value="#ff0000" table:operator="=" table:data-type="text-color"`)}</table:filter>`],
    ["an OR across fields", `<table:filter><table:filter-or>${condition(0, `table:value="East" table:operator="="`)}${condition(1, `table:value="5" table:operator="=" table:data-type="number"`)}</table:filter-or></table:filter>`],
    ["an or-ed field inside an OR across fields", `<table:filter><table:filter-and><table:filter-or>${condition(0, `table:value="a" table:operator="="`)}${condition(2, `table:value="b" table:operator="="`)}</table:filter-or></table:filter-and></table:filter>`],
    ["a field outside the range", `<table:filter>${condition(4, `table:value="East" table:operator="="`)}</table:filter>`],
    ["three conditions on a field", `<table:filter><table:filter-and>${[1, 2, 3].map((n) => condition(1, `table:value="${n}" table:operator="!=" table:data-type="number"`)).join("")}</table:filter-and></table:filter>`],
    ["a values list and a condition on one field", `<table:filter><table:filter-and>${condition(0, `table:value="" table:operator="empty"`)}${condition(0, `table:value="a" table:operator="="`)}</table:filter-and></table:filter>`],
    ["a nested filter-and", `<table:filter><table:filter-and><table:filter-and>${condition(0, `table:value="a" table:operator="="`)}</table:filter-and></table:filter-and></table:filter>`],
    ["a criteria-range source (an advanced filter)", `<table:filter table:condition-source-range-address="Sheet1.F1:Sheet1.F3">${condition(0, `table:value="a" table:operator="="`)}</table:filter>`],
    ["a copy-output range", `<table:filter table:target-range-address="Sheet1.H1">${condition(0, `table:value="a" table:operator="="`)}</table:filter>`],
    ["duplicates hidden", `<table:filter table:display-duplicates="false">${condition(0, `table:value="a" table:operator="="`)}</table:filter>`],
    ["a top10 count that is not positive", `<table:filter>${condition(1, `table:value="0" table:operator="top values" table:data-type="number"`)}</table:filter>`],
  ];

  it.each(UNMAPPED)("%s", (_label, wire) => {
    const autoFilter = readOne(anonymous(wire))!;
    expect(autoFilter.columns).toEqual([]);
    expect(autoFilter.odsUnmappedFilter).toBe(wire);
    // Written back verbatim, whatever the columns say.
    expect(odsFilterXml({ ...autoFilter, columns: [{ colId: 0, criteria: values(["x"]) }] })).toBe(wire);
  });
});

describe("criteria with no ODS form", () => {
  it("leaves out preserved, dynamic, date groups and an empty values list; keeps a values list's values without its groups", () => {
    const groups = [{ grouping: "month" as const, year: 2026, month: 3 }];
    const none: FilterCriteria[] = [
      { kind: "preserved", xml: `<colorFilter dxfId="0"/>`, label: "colour" },
      { kind: "dynamic", type: "thisMonth" },
      { kind: "values", values: [], blanks: false, dateGroups: groups },
      { kind: "values", values: [], blanks: false, dateGroups: [] },
    ];
    for (const criteria of none) {
      expect(filterOf([{ colId: 0, criteria }])).toBe("");
    }
    expect(filterOf([{ colId: 0, criteria: { kind: "dynamic", type: "aboveAverage" } }, { colId: 1, criteria: custom("and", ["greaterThan", "1"]) }])).toBe(
      `<table:filter>${condition(1, `table:value="1" table:operator="&gt;" table:data-type="number"`)}</table:filter>`,
    );
    expect(filterOf([{ colId: 0, criteria: { kind: "values", values: ["East"], blanks: false, dateGroups: groups } }])).toBe(
      `<table:filter><table:filter-condition table:field-number="0" table:value="East" table:operator="="><table:filter-set-item table:value="East"/></table:filter-condition></table:filter>`,
    );
    // A column with no criteria (a hidden button) writes nothing; one outside the range neither.
    expect(filterOf([{ colId: 0, hideButton: true }, { colId: 9, criteria: values(["x"]) }])).toBe("");
  });
});

describe("the sort record", () => {
  const SORT = `<table:sort table:case-sensitive="true"><table:sort-by table:field-number="1" table:data-type="automatic" table:order="descending"/><table:sort-by table:field-number="0" table:data-type="automatic" table:order="ascending"/></table:sort>`;

  it("reads table:sort into autoFilter.sort — the data rows, absolute key columns, case sensitivity, the element verbatim — and writes it back", () => {
    const element = `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.C3:Sheet1.F9" table:display-filter-buttons="true">${SORT}</table:database-range>`;
    const autoFilter = readOne(element)!;
    expect(autoFilter.range).toEqual({ top: 2, left: 2, bottom: 8, right: 5 });
    expect(autoFilter.sort).toEqual({
      ref: { top: 3, left: 2, bottom: 8, right: 5 },
      keys: [
        { col: 3, descending: true },
        { col: 2, descending: false },
      ],
      caseSensitive: true,
      sourceXml: SORT,
    });
    // Unchanged: verbatim. A new sort (no sourceXml): generated, the same here.
    expect(odsSortXml(autoFilter)).toBe(SORT);
    const { sourceXml: _source, ...fresh } = autoFilter.sort!;
    expect(odsSortXml({ ...autoFilter, sort: fresh })).toBe(SORT);
    expect(odsSortXml({ ...autoFilter, sort: { ref: fresh.ref, keys: [{ col: 4, descending: false }, { col: 9, descending: true }] } })).toBe(
      `<table:sort><table:sort-by table:field-number="2" table:data-type="automatic" table:order="ascending"/></table:sort>`,
    );
    expect(odsSortXml({ ...autoFilter, sort: { ref: fresh.ref, keys: [] } })).toBe("");
  });

  it("a user-list key or one outside the range is no key; an xlsx sortState is never written as ODS", () => {
    const sort = `<table:sort><table:sort-by table:field-number="0" table:data-type="UserList1"/><table:sort-by table:field-number="7"/><table:sort-by table:field-number="2"/></table:sort>`;
    const autoFilter = readOne(anonymous(sort))!;
    expect(autoFilter.sort).toEqual({ ref: { top: 1, left: 0, bottom: 4, right: 3 }, keys: [{ col: 2, descending: false }], sourceXml: sort });
    expect(odsSortXml({ ...autoFilter, sort: { ...autoFilter.sort!, sourceXml: `<sortState ref="A2:D5"/>` } })).toBe(
      `<table:sort><table:sort-by table:field-number="2" table:data-type="automatic" table:order="ascending"/></table:sort>`,
    );
  });
});

const named = (name: string, target: string, buttons = true): string =>
  `<table:database-range table:name="${name}" table:target-range-address="${target}"${buttons ? ` table:display-filter-buttons="true"` : ""}/>`;

describe("the claim: which database range is a sheet's autofilter", () => {
  it("takes the anonymous range first, a named one only when the sheet has none, and needs filter buttons", () => {
    const suffix = databaseRanges(
      named("Sales", "Sheet1.A1:Sheet1.B9"),
      named("__Anonymous_Sheet_DB__0", "Sheet1.A1:Sheet1.D5"),
      named("Other", "Sheet2.B2:Sheet2.C4"),
      named("Plain", "Sheet3.A1:Sheet3.B2", false),
      named("Later", "Sheet2.E1:Sheet2.F2"),
    );
    const filters = readOdsAutoFilters(suffix, ["Sheet1", "Sheet2", "Sheet3"]);
    expect([...filters.keys()]).toEqual([0, 1]);
    expect(filters.get(0)).toMatchObject({ range: RANGE, odsName: "__Anonymous_Sheet_DB__0", sourceXml: named("__Anonymous_Sheet_DB__0", "Sheet1.A1:Sheet1.D5") });
    expect(filters.get(1)).toMatchObject({ range: { top: 1, left: 1, bottom: 3, right: 2 }, odsName: "Other" });
    // The claimed elements' spans, for the retained-reference pass.
    const spans = odsOwnedDatabaseRangeSpans(suffix, ["Sheet1", "Sheet2", "Sheet3"]).map((span) => suffix.slice(span.start, span.end));
    expect(spans).toEqual([named("__Anonymous_Sheet_DB__0", "Sheet1.A1:Sheet1.D5"), named("Other", "Sheet2.B2:Sheet2.C4")]);
  });

  it("reads quoted names, `$` and either end order; refuses another sheet, 3-D, lists, whole columns, unparseable and malformed elements", () => {
    expect(readOdsAutoFilters(databaseRanges(named("__Anonymous_Sheet_DB__0", "$'Q1 ''Data'''.$D$5:.$A$1")), ["Q1 'Data'"]).get(0)!.range).toEqual(RANGE);
    expect(readOdsAutoFilters(databaseRanges(named("x", "sheet1.A1:SHEET1.D5")), ["Sheet1"]).get(0)!.range).toEqual(RANGE);
    for (const target of ["Sheet9.A1:Sheet9.D5", "Sheet1.A1:Sheet2.D5", "Sheet1.A1:Sheet1.B2 Sheet1.D1:Sheet1.D2", "Sheet1.A:Sheet1.D", "$#REF!.A1:.D5", ".A1:.D5", ""]) {
      expect(readOdsAutoFilters(databaseRanges(named("__Anonymous_Sheet_DB__0", target)), ["Sheet1", "Sheet2"]).size, target).toBe(0);
    }
    const malformed = `${ANONYMOUS_HEAD}<table:filter><table:filter-condition></table:filter></table:database-range>`;
    expect(readOdsAutoFilters(databaseRanges(malformed), ["Sheet1"]).size).toBe(0);
    // A malformed anonymous range does not hide a well-formed named one.
    expect(readOdsAutoFilters(databaseRanges(malformed, named("Sales", "Sheet1.A1:Sheet1.D5")), ["Sheet1"]).get(0)!.odsName).toBe("Sales");
    expect(odsAutoFilterClaims(odsDatabaseRanges(`<table:named-expressions/>`), ["Sheet1"])).toEqual([]);
  });

  it("writes a target as LibreOffice does, the sheet on both ends", () => {
    expect(odsTargetRangeAddress("Sheet1", RANGE)).toBe("Sheet1.A1:Sheet1.D5");
    expect(odsTargetRangeAddress("Q1 'Data'", { top: 9, left: 26, bottom: 19, right: 27 })).toBe("'Q1 ''Data'''.AA10:'Q1 ''Data'''.AB20");
    expect(odsTargetRangeAddress("v1.2", RANGE)).toBe("'v1.2'.A1:'v1.2'.D5");
    expect(odsDatabaseRangeXml({ range: RANGE, columns: [] }, "Sheet1", "__Anonymous_Sheet_DB__3")).toBe(
      `<table:database-range table:name="__Anonymous_Sheet_DB__3" table:target-range-address="Sheet1.A1:Sheet1.D5" table:display-filter-buttons="true"/>`,
    );
  });
});

// --- a synthetic LibreOffice-shaped workbook --------------------------------------------------

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" ` +
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"`;
const MANIFEST =
  `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
  `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
  `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
  `</manifest:manifest>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"/></office:styles></office:document-styles>`;

const str = (text: string): string => `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
const num = (value: number): string => `<table:table-cell office:value-type="float" office:value="${value}" calcext:value-type="float"><text:p>${value}</text:p></table:table-cell>`;
const row = (cells: string, attrs = ""): string => `<table:table-row${attrs}>${cells}<table:table-cell table:number-columns-repeated="2"/></table:table-row>`;
const TAIL_ROWS = `<table:table-row table:number-rows-repeated="1048571"><table:table-cell table:number-columns-repeated="4"/></table:table-row>`;
/** Region / Amount: West filtered out (East and blanks pass), the blank one hidden by the user. */
const SALES =
  row(str("Region") + str("Amount")) +
  row(str("East") + num(10)) +
  row(str("West") + num(20), ` table:visibility="filter"`) +
  row(str("East") + num(30)) +
  row(`<table:table-cell/>` + num(40), ` table:visibility="collapse"`) +
  TAIL_ROWS;
const table = (name: string, rows: string): string => `<table:table table:name="${name}"><table:table-column table:number-columns-repeated="4"/>${rows}</table:table>`;
const OTHER_ROWS = row(str("x") + num(1)) + `<table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="4"/></table:table-row>`;

const EAST_OR_BLANK = `<table:filter><table:filter-condition table:field-number="0" table:value="East" table:operator="="><table:filter-set-item table:value="East"/><table:filter-set-item table:value=""/></table:filter-condition></table:filter>`;
const SALES_FILTER = `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.B5" table:display-filter-buttons="true">${EAST_OR_BLANK}</table:database-range>`;
const DB1 = `<table:database-range table:name="db1" table:target-range-address="Sheet2.A1:Sheet2.B3"/>`;

interface Book {
  tables?: string;
  /** Everything between the last table and `</office:spreadsheet>`. */
  suffix?: string;
}

async function bookBytes(book: Book): Promise<{ bytes: Uint8Array; parts: Map<string, Uint8Array> }> {
  const content =
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles/><office:body><office:spreadsheet>` +
    `${book.tables ?? table("Sheet1", SALES) + table("Sheet2", OTHER_ROWS)}${book.suffix ?? `<table:named-expressions/>${databaseRanges(DB1, SALES_FILTER)}`}` +
    `</office:spreadsheet></office:body></office:document-content>`;
  const bytes = await zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    { name: "META-INF/manifest.xml", data: encoder.encode(MANIFEST) },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(STYLES) },
  ]);
  return { bytes, parts: await unzip(bytes) };
}

async function open(book: Book = {}): Promise<{ source: Map<string, Uint8Array>; model: SpreadsheetModel }> {
  const { bytes, parts } = await bookBytes(book);
  return { source: parts, model: await readSpreadsheetBytes(bytes, "ods", "book") };
}

/** The saved package's content.xml after its last table. */
async function savedSuffix(model: SpreadsheetModel): Promise<string> {
  const bytes = await writeOds(model);
  expect(await packageIntegrity(bytes)).toEqual([]);
  const content = decoder.decode((await unzip(bytes)).get("content.xml"));
  return content.slice(content.lastIndexOf("</table:table>") + "</table:table>".length, content.indexOf("</office:spreadsheet>"));
}

async function reopen(model: SpreadsheetModel): Promise<SpreadsheetModel> {
  return readSpreadsheetBytes(await writeOds(model), "ods", "reopened");
}

/** An autofilter without what the reader records about its source. */
function modelled(autoFilter: AutoFilter | undefined): AutoFilter | undefined {
  if (!autoFilter) {
    return undefined;
  }
  const { sourceXml: _source, odsName: _name, ...rest } = autoFilter;
  if (rest.sort) {
    const { sourceXml: _sort, ...sort } = rest.sort;
    rest.sort = sort;
  }
  return rest;
}

const visibilityOf = (sheet: Sheet): [number, string][] =>
  sheet.rows.filter((each) => each.hidden || each.filterHidden).map((each) => [each.index, `${each.hidden ? "hidden" : ""}${each.filterHidden ? "filterHidden" : ""}`]);

/** The sheet at `index` with its autofilter replaced (as a command would) and marked dirty. */
function setAutoFilter(model: SpreadsheetModel, index: number, autoFilter: AutoFilter | undefined): void {
  const { autoFilter: _old, ...rest } = model.sheets[index]!;
  model.sheets[index] = autoFilter ? { ...rest, autoFilter, dirty: true } : { ...rest, dirty: true };
}

describe("reading a workbook", () => {
  it("gives Sheet1 its autofilter, and records every row as table:visibility says — never classified", async () => {
    const { model } = await open();
    expect(model.sheets[0]!.autoFilter).toEqual({
      range: { top: 0, left: 0, bottom: 4, right: 1 },
      columns: [{ colId: 0, criteria: values(["East"], true) }],
      sourceXml: SALES_FILTER,
      odsName: "__Anonymous_Sheet_DB__0",
    });
    expect(model.sheets[1]!.autoFilter).toBeUndefined();
    // West fails and is filter-hidden; the blank row passes but the user hid it.
    expect(visibilityOf(model.sheets[0]!)).toEqual([
      [2, "filterHidden"],
      [4, "hidden"],
    ]);
  });

  it("reads a filter-hidden blank row as a Row with no cells, and keeps a visible row that fails visible", async () => {
    const rows = row(str("Region")) + row(str("West")) + `<table:table-row table:visibility="filter"><table:table-cell table:number-columns-repeated="4"/></table:table-row>` + TAIL_ROWS;
    const filter = `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.A3" table:display-filter-buttons="true">${EAST_OR_BLANK}</table:database-range>`;
    const parsed = await parseOds(new Uint8Array(0), await bookBytes({ tables: table("Sheet1", rows), suffix: databaseRanges(filter) }).then((book) => book.parts));
    const sheet = parsed.sheets[0]!;
    expect(sheet.rows.find((each) => each.index === 1)).toEqual({ index: 1, cells: [{ col: 0, value: "West" }] });
    expect(sheet.rows.find((each) => each.index === 2)).toEqual({ index: 2, cells: [], filterHidden: true });
  });
});

describe("writing a workbook", () => {
  it("downloads an untouched filtered workbook byte-identical, part for part", async () => {
    const { source, model } = await open();
    const out = await unzip(await writeOds(model));
    expect([...out.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("re-emits an unchanged autofilter's element verbatim inside a rebuilt content.xml", async () => {
    const { model } = await open();
    model.sheets[0]!.dirty = true;
    expect(await savedSuffix(model)).toBe(`<table:named-expressions/>${databaseRanges(DB1, SALES_FILTER)}`);
  });

  it("regenerates the element in place after a criteria change, the other database ranges verbatim; write → read reproduces it", async () => {
    const { model } = await open();
    const autoFilter: AutoFilter = { range: model.sheets[0]!.autoFilter!.range, columns: [{ colId: 1, criteria: custom("and", ["greaterThan", "15"]) }] };
    setAutoFilter(model, 0, autoFilter);
    expect(await savedSuffix(model)).toBe(
      `<table:named-expressions/>${databaseRanges(
        DB1,
        `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.B5" table:display-filter-buttons="true">` +
          `<table:filter>${condition(1, `table:value="15" table:operator="&gt;" table:data-type="number"`)}</table:filter></table:database-range>`,
      )}`,
    );
    const reread = await reopen(model);
    expect(modelled(reread.sheets[0]!.autoFilter)).toEqual(autoFilter);
    expect(reread.sheets[0]!.autoFilter!.odsName).toBe("__Anonymous_Sheet_DB__0");
    expect(visibilityOf(reread.sheets[0]!)).toEqual(visibilityOf(model.sheets[0]!));
  });

  it("patches a claimed named range: its name, unmodelled attributes and children kept, the filter and sort regenerated in ODF order", async () => {
    const element =
      `<table:database-range table:name="Sales" table:target-range-address="Sheet1.A1:Sheet1.B5" table:display-filter-buttons="true" table:refresh-delay="PT1M">` +
      `<table:database-source-table table:database-name="db" table:database-table-name="t"/>${EAST_OR_BLANK}<table:subtotal-rules/></table:database-range>`;
    const { model } = await open({ suffix: databaseRanges(element) });
    expect(model.sheets[0]!.autoFilter!.odsName).toBe("Sales");
    const range = { top: 0, left: 0, bottom: 6, right: 1 };
    setAutoFilter(model, 0, { range, columns: [], sort: { ref: { ...range, top: 1 }, keys: [{ col: 1, descending: true }] } });
    expect(await savedSuffix(model)).toBe(
      databaseRanges(
        `<table:database-range table:name="Sales" table:target-range-address="Sheet1.A1:Sheet1.B7" table:display-filter-buttons="true" table:refresh-delay="PT1M">` +
          `<table:database-source-table table:database-name="db" table:database-table-name="t"/>` +
          `<table:sort><table:sort-by table:field-number="1" table:data-type="automatic" table:order="descending"/></table:sort><table:subtotal-rules/></table:database-range>`,
      ),
    );
  });

  it("appends a new sheet's filter as its own anonymous range", async () => {
    const { model } = await open();
    const autoFilter: AutoFilter = { range: { top: 0, left: 0, bottom: 0, right: 1 }, columns: [{ colId: 0, criteria: { kind: "top10", top: false, percent: true, value: 50 } }] };
    setAutoFilter(model, 1, autoFilter);
    expect(await savedSuffix(model)).toBe(
      `<table:named-expressions/>${databaseRanges(
        DB1,
        SALES_FILTER,
        `<table:database-range table:name="__Anonymous_Sheet_DB__1" table:target-range-address="Sheet2.A1:Sheet2.B1" table:display-filter-buttons="true">` +
          `<table:filter>${condition(0, `table:value="50" table:operator="bottom percent" table:data-type="number"`)}</table:filter></table:database-range>`,
      )}`,
    );
    expect(modelled((await reopen(model)).sheets[1]!.autoFilter)).toEqual(autoFilter);
  });

  it("creates a missing table:database-ranges after the named expressions, before the pilot tables", async () => {
    const pilots = `<table:data-pilot-tables><table:data-pilot-table table:name="p" table:target-range-address="Sheet2.D1:Sheet2.E2"/></table:data-pilot-tables>`;
    const { model } = await open({ suffix: `<table:named-expressions/>${pilots}` });
    setAutoFilter(model, 1, { range: { top: 0, left: 0, bottom: 1, right: 1 }, columns: [] });
    const added = `<table:database-ranges><table:database-range table:name="__Anonymous_Sheet_DB__1" table:target-range-address="Sheet2.A1:Sheet2.B2" table:display-filter-buttons="true"/></table:database-ranges>`;
    expect(await savedSuffix(model)).toBe(`<table:named-expressions/>${added}${pilots}`);
    // With nothing after the tables: before </office:spreadsheet>.
    const bare = await open({ suffix: "" });
    setAutoFilter(bare.model, 1, { range: { top: 0, left: 0, bottom: 1, right: 1 }, columns: [] });
    expect(await savedSuffix(bare.model)).toBe(added);
  });

  it("removes the element of an autofilter the model no longer has, and an emptied container", async () => {
    const { model } = await open();
    setAutoFilter(model, 0, undefined);
    expect(await savedSuffix(model)).toBe(`<table:named-expressions/>${databaseRanges(DB1)}`);
    const alone = await open({ suffix: `<table:named-expressions/>${databaseRanges(SALES_FILTER)}` });
    setAutoFilter(alone.model, 0, undefined);
    expect(await savedSuffix(alone.model)).toBe(`<table:named-expressions/>`);
    expect((await reopen(alone.model)).sheets[0]!.autoFilter).toBeUndefined();
  });

  it("moves the range once on a structural edit — through the model, never the retained text — and writes it", async () => {
    const { model } = await open();
    const sheetId = model.sheets[0]!.id;
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId, at: 0, count: 2 })).toEqual({ ok: true });
    expect(applyStructuralEdit(model, { kind: "insertCols", sheetId, at: 0, count: 1 })).toEqual({ ok: true });
    // The retained element is untouched: only the model moved.
    expect(model.odsContentEnvelope!.suffix).toContain(SALES_FILTER);
    expect(model.sheets[0]!.autoFilter!.range).toEqual({ top: 2, left: 1, bottom: 6, right: 2 });
    expect(await savedSuffix(model)).toBe(
      `<table:named-expressions/>${databaseRanges(DB1, SALES_FILTER.replace(`"Sheet1.A1:Sheet1.B5"`, `"Sheet1.B3:Sheet1.C7"`))}`,
    );
    const reread = await reopen(model);
    expect(modelled(reread.sheets[0]!.autoFilter)).toEqual(modelled(model.sheets[0]!.autoFilter));
    expect(visibilityOf(reread.sheets[0]!)).toEqual([
      [4, "filterHidden"],
      [6, "hidden"],
    ]);
  });

  it("moves a verbatim filter's criteria range on another sheet once — the model's copy, by a structural edit and by a move — and writes it", async () => {
    const criteria = `<table:filter table:condition-source-range-address="Sheet2.A1:Sheet2.B2">${condition(0, `table:value="East" table:operator="="`)}</table:filter>`;
    const ranged = `<table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.B5" table:display-filter-buttons="true">${criteria}</table:database-range>`;
    const { model } = await open({ suffix: `<table:named-expressions/>${databaseRanges(ranged)}` });
    expect(model.sheets[0]!.autoFilter!.odsUnmappedFilter).toBe(criteria);
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: model.sheets[1]!.id, at: 0, count: 2 })).toEqual({ ok: true });
    const inserted = criteria.replace(`"Sheet2.A1:Sheet2.B2"`, `"Sheet2.A3:Sheet2.B4"`);
    expect(model.sheets[0]!.autoFilter!.odsUnmappedFilter).toBe(inserted);
    expect(model.sheets[0]!.dirty).toBe(true);
    // The retained element is untouched: only the model moved it.
    expect(model.odsContentEnvelope!.suffix).toContain(ranged);
    expect(await savedSuffix(model)).toBe(`<table:named-expressions/>${databaseRanges(ranged.replace(criteria, inserted))}`);
    expect(moveRange(model, { fromSheetId: model.sheets[1]!.id, from: { top: 2, left: 0, bottom: 3, right: 1 }, toSheetId: model.sheets[1]!.id, toTop: 5, toLeft: 2 })).toEqual({ ok: true });
    const moved = criteria.replace(`"Sheet2.A1:Sheet2.B2"`, `"Sheet2.C6:Sheet2.D7"`);
    expect(model.sheets[0]!.autoFilter!.odsUnmappedFilter).toBe(moved);
    expect(await savedSuffix(model)).toBe(`<table:named-expressions/>${databaseRanges(ranged.replace(criteria, moved))}`);
    expect((await reopen(model)).sheets[0]!.autoFilter!.odsUnmappedFilter).toBe(moved);
  });

  it("re-serialises a verbatim element's target when its sheet is renamed", async () => {
    const { model } = await open();
    renameSheet(model, 0, "Q1 Sales");
    expect(await savedSuffix(model)).toBe(
      `<table:named-expressions/>${databaseRanges(DB1, SALES_FILTER.replace(`"Sheet1.A1:Sheet1.B5"`, `"'Q1 Sales'.A1:'Q1 Sales'.B5"`))}`,
    );
    const reread = await reopen(model);
    expect(reread.sheets[0]!.name).toBe("Q1 Sales");
    expect(modelled(reread.sheets[0]!.autoFilter)).toEqual(modelled(model.sheets[0]!.autoFilter));
  });

  it("renumbers an anonymous range whose sheet moved, and drops the ranges of a deleted sheet", async () => {
    const filter2 = `<table:database-range table:name="__Anonymous_Sheet_DB__1" table:target-range-address="Sheet2.A1:Sheet2.B1" table:display-filter-buttons="true"/>`;
    const { model } = await open({ suffix: databaseRanges(SALES_FILTER, filter2) });
    deleteSheet(model, 0);
    expect(await savedSuffix(model)).toBe(databaseRanges(filter2.replace("__Anonymous_Sheet_DB__1", "__Anonymous_Sheet_DB__0")));
  });

  it("writes a duplicated sheet's autofilter as its own anonymous range, the origin's kept", async () => {
    const { model } = await open();
    duplicateSheet(model, 0);
    expect(model.sheets[1]!.autoFilter?.odsName).toBeUndefined();
    expect(await savedSuffix(model)).toBe(
      `<table:named-expressions/>${databaseRanges(
        DB1,
        SALES_FILTER,
        `<table:database-range table:name="__Anonymous_Sheet_DB__1" table:target-range-address="'Sheet1 (2)'.A1:'Sheet1 (2)'.B5" table:display-filter-buttons="true">${EAST_OR_BLANK}</table:database-range>`,
      )}`,
    );
    const reread = await reopen(model);
    expect(modelled(reread.sheets[1]!.autoFilter)).toEqual(modelled(reread.sheets[0]!.autoFilter));
    expect(visibilityOf(reread.sheets[1]!)).toEqual(visibilityOf(reread.sheets[0]!));
  });

  it("writes every sheet's autofilter from scratch, and reads it back", async () => {
    const model = createBlankWorkbook();
    const rows: Row[] = [
      { index: 0, cells: [{ col: 0, value: "Region" }] },
      { index: 1, cells: [{ col: 0, value: "East" }] },
      { index: 2, cells: [{ col: 0, value: "West" }], filterHidden: true },
      { index: 3, cells: [], filterHidden: true },
      { index: 4, cells: [{ col: 0, value: "East" }], hidden: true },
    ];
    const autoFilter: AutoFilter = {
      range: { top: 0, left: 0, bottom: 4, right: 0 },
      columns: [{ colId: 0, criteria: values(["East"]) }],
      sort: { ref: { top: 1, left: 0, bottom: 4, right: 0 }, keys: [{ col: 0, descending: false }], caseSensitive: true },
    };
    model.sheets[0] = { ...model.sheets[0]!, rows, autoFilter };
    const bytes = await writeOds(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const content = decoder.decode((await unzip(bytes)).get("content.xml"));
    expect(content).toContain(
      `</table:table><table:database-ranges><table:database-range table:name="__Anonymous_Sheet_DB__0" table:target-range-address="Sheet1.A1:Sheet1.A5" table:display-filter-buttons="true">` +
        `<table:filter><table:filter-condition table:field-number="0" table:value="East" table:operator="="><table:filter-set-item table:value="East"/></table:filter-condition></table:filter>` +
        `<table:sort table:case-sensitive="true"><table:sort-by table:field-number="0" table:data-type="automatic" table:order="ascending"/></table:sort>` +
        `</table:database-range></table:database-ranges></office:spreadsheet>`,
    );
    const reread = await readSpreadsheetBytes(bytes, "ods", "reopened");
    expect(modelled(reread.sheets[0]!.autoFilter)).toEqual(autoFilter);
    expect(visibilityOf(reread.sheets[0]!)).toEqual([
      [2, "filterHidden"],
      [3, "filterHidden"],
      [4, "hidden"],
    ]);
  });
});

const sheetOf = (rows: Row[]): Sheet => ({ id: "s", name: "S", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] });
/** serializeTable for a new sheet holding `rows`. */
const writeRows = (rows: Row[]): string =>
  serializeTable(sheetOf(rows), { model: { styles: [{}], dateSystem: "1900", format: "ods" }, registry: new OdsStyleRegistry(), locale: SPREADSHEET_LOCALES["en-US"] });
const tableRow = (attrs: string): string => `<table:table-row table:style-name="de-ro1"${attrs}><table:table-cell/></table:table-row>`;

describe("row visibility in the table writer", () => {
  it("writes collapse for a hidden row whatever filterHidden is, filter for a filter-hidden one; visibility keeps rows apart in the repeat compression", () => {
    const xml = writeRows([
      { index: 0, cells: [], filterHidden: true },
      { index: 1, cells: [], filterHidden: true },
      { index: 2, cells: [], hidden: true, filterHidden: true },
      { index: 3, cells: [], hidden: true },
      { index: 4, cells: [] },
      { index: 5, cells: [], filterHidden: true },
    ]);
    expect(xml).toBe(
      `<table:table table:name="S" table:style-name="de-ta1"><table:table-column table:style-name="de-co1" table:default-cell-style-name="Default"/>` +
        tableRow(` table:number-rows-repeated="2" table:visibility="filter"`) +
        tableRow(` table:number-rows-repeated="2" table:visibility="collapse"`) +
        tableRow("") +
        tableRow(` table:visibility="filter"`) +
        `</table:table>`,
    );
  });
});
