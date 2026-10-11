// SPEC-128 "Criteria semantics", "Filter-hidden or user-hidden", "Choosing
// the range": every criterion and Excel encoding, the date dynamics under
// setClock (never the real today), applying (AND across columns, the used
// range, growth), reapply with a preserved criterion, the read
// classification, the commands on a sheet holding both kinds of hidden row,
// visibleRowSpans, the popover's value list and what the UI reads.
import { afterEach, describe, expect, it } from "vitest";
import { setClock } from "../core/clock";
import { findRow, type GridRange } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import {
  applyAutoFilter,
  autoFilterRangeFor,
  classifyHiddenRows,
  clearFilters,
  columnValueKind,
  EMPTY_FILTER_REGION_MESSAGE,
  FILTER_VALUE_LIST_MAX,
  filterButtonState,
  filterRecordCounts,
  filterValueList,
  hasCriteria,
  reapplyFilter,
  rowPasses,
  setFilterCriteria,
  toggleAutoFilter,
  visibleRowSpans,
  type AutoFilter,
  type FilterCriteria,
  type FilterOperator,
} from "./filter";
import { buildMergeIndex } from "./merges";
import type { Cell, CellValue, Row, Sheet, SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import type { NavContext, SheetSelection } from "./selection";
import { dateToSerial } from "./serial";

const DATE_STYLE = 1;
const PERCENT_STYLE = 2;

function sheet(fields: Partial<Sheet> = {}): Sheet {
  return { id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(s: Sheet): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "book",
    title: "T",
    format: "xlsx",
    readerVersion: 9,
    sheets: [s],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}, { numberFormat: "yyyy-mm-dd" }, { numberFormat: "0%" }],
    defaultStyleId: 0,
  };
}

const ctx = workbookFormatContext({ dateSystem: "1900", format: "xlsx" }, SPREADSHEET_LOCALES["en-US"]);
const cell = (col: number, value: CellValue, extra: Partial<Cell> = {}): Cell => ({ col, value, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const date = (y: number, m: number, d: number): number => dateToSerial(y, m, d, "1900");

/** One-column list: header "H" at row 0, then `values` from row 1, under
 * an autofilter over the whole list with `criteria` on column 0. */
function list(values: CellValue[], criteria?: FilterCriteria, styleId?: number): SpreadsheetModel {
  const rows = [row(0, [cell(0, "H")]), ...values.map((v, i) => row(i + 1, v === null ? [] : [cell(0, v, styleId === undefined ? {} : { styleId })]))];
  const autoFilter: AutoFilter = { range: range(0, 0, values.length, 0), columns: criteria ? [{ colId: 0, criteria }] : [] };
  return model(sheet({ rows: rows.filter((r) => r.cells.length > 0), autoFilter }));
}

/** The values (rows 1..n) a criterion passes. */
function passing(values: CellValue[], criteria: FilterCriteria, styleId?: number): CellValue[] {
  const m = list(values, criteria, styleId);
  return values.filter((_, i) => rowPasses(m, m.sheets[0]!, i + 1, ctx));
}

const hiddenRows = (m: SpreadsheetModel, flag: "hidden" | "filterHidden"): number[] => m.sheets[0]!.rows.filter((r) => r[flag]).map((r) => r.index);

afterEach(() => {
  setClock(null);
});

describe("values", () => {
  const values = ["East", "west", null, "North", 12];

  it("passes a display text equal to an entry, case-insensitively, and a blank only with (Blanks)", () => {
    expect(passing(values, { kind: "values", values: ["WEST", "12"], blanks: false, dateGroups: [] })).toEqual(["west", 12]);
    expect(passing(values, { kind: "values", values: ["East"], blanks: true, dateGroups: [] })).toEqual(["East", null]);
  });

  it("compares the display text, not the value", () => {
    expect(passing([0.5, 0.25], { kind: "values", values: ["50%"], blanks: false, dateGroups: [] }, PERCENT_STYLE)).toEqual([0.5]);
  });

  it("nothing listed hides every data row", () => {
    expect(passing(values, { kind: "values", values: [], blanks: false, dateGroups: [] })).toEqual([]);
  });

  it("date groups match a date cell's parts down to their grouping", () => {
    const dates = [date(2026, 3, 4), date(2026, 3, 9), date(2026, 4, 4), date(2025, 3, 4)];
    expect(passing(dates, { kind: "values", values: [], blanks: false, dateGroups: [{ grouping: "month", year: 2026, month: 3 }] }, DATE_STYLE)).toEqual(dates.slice(0, 2));
    expect(passing(dates, { kind: "values", values: [], blanks: false, dateGroups: [{ grouping: "year", year: 2026 }] }, DATE_STYLE)).toEqual(dates.slice(0, 3));
    expect(passing(dates, { kind: "values", values: [], blanks: false, dateGroups: [{ grouping: "day", year: 2026, month: 3, day: 9 }] }, DATE_STYLE)).toEqual([dates[1]]);
    // A number that is not a date cell never matches a date group.
    expect(passing([date(2026, 3, 4)], { kind: "values", values: [], blanks: false, dateGroups: [{ grouping: "year", year: 2026 }] })).toEqual([]);
  });
});

/** A custom criterion of one condition. */
function one(operator: FilterOperator, value: string): FilterCriteria {
  return { kind: "custom", join: "and", conditions: [{ operator, value }] };
}

describe("custom", () => {
  it("notEqual \" \" is non-blank (Excel's encoding)", () => {
    expect(passing(["a", null, 0, ""], one("notEqual", " "))).toEqual(["a", 0]);
  });

  it("wildcards * ? under equal/notEqual match the display text, ~ escapes", () => {
    const values = ["Apple pie", "apricot", "grape", "a*b", "ab"];
    expect(passing(values, one("equal", "ap*"))).toEqual(["Apple pie", "apricot"]);
    expect(passing(values, one("equal", "*pe"))).toEqual(["grape"]);
    expect(passing(values, one("equal", "*ric*"))).toEqual(["apricot"]);
    expect(passing(values, one("equal", "a?"))).toEqual(["ab"]);
    expect(passing(values, one("equal", "a~*b"))).toEqual(["a*b"]);
    expect(passing(values, one("notEqual", "*p*"))).toEqual(["a*b", "ab"]);
  });

  it("an invariant number against a numeric cell compares numerically", () => {
    expect(passing([5, 50, 500], one("greaterThan", "10"))).toEqual([50, 500]);
    // A text cell compares its display text through the collator instead.
    expect(passing(["50", "05"], one("greaterThan", "10"))).toEqual(["50"]);
    expect(passing([5, 50, 500], one("lessThanOrEqual", "50"))).toEqual([5, 50]);
    expect(passing([5, 50, 500], one("equal", "50"))).toEqual([50]);
    expect(passing([5, 50, 500], one("notEqual", "50"))).toEqual([5, 500]);
    expect(passing([5, 50, 500], one("greaterThanOrEqual", "5E1"))).toEqual([50, 500]);
  });

  it("otherwise text compares through the collator; numeric cells never match a text ordering", () => {
    expect(passing(["apple", "Banana", "cherry", 7], one("greaterThan", "b"))).toEqual(["Banana", "cherry"]);
    expect(passing(["apple", "Banana", "cherry"], one("lessThan", "B"))).toEqual(["apple"]);
    expect(passing(["Apple", "pear"], one("equal", "apple"))).toEqual(["Apple"]);
  });

  it("a blank fails every operator except notEqual", () => {
    for (const operator of ["equal", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"] as const) {
      expect(passing([null], one(operator, "0"))).toEqual([]);
    }
    expect(passing([null], one("notEqual", "x"))).toEqual([null]);
  });

  it("joins two conditions with and/or", () => {
    const between: FilterCriteria = { kind: "custom", join: "and", conditions: [{ operator: "greaterThanOrEqual", value: "100" }, { operator: "lessThanOrEqual", value: "900" }] };
    expect(passing([50, 100, 500, 900, 950], between)).toEqual([100, 500, 900]);
    const outside: FilterCriteria = { kind: "custom", join: "or", conditions: [{ operator: "lessThan", value: "100" }, { operator: "greaterThan", value: "900" }] };
    expect(passing([50, 100, 500, 900, 950], outside)).toEqual([50, 950]);
  });
});

describe("top10", () => {
  const values = [10, 40, "x", 20, 50, 30, null];

  it("ranks the numeric cells of every data row: top/bottom items", () => {
    expect(passing(values, { kind: "top10", top: true, percent: false, value: 2 })).toEqual([40, 50]);
    expect(passing(values, { kind: "top10", top: false, percent: false, value: 2 })).toEqual([10, 20]);
    expect(passing(values, { kind: "top10", top: true, percent: false, value: 10 })).toEqual([10, 40, 20, 50, 30]);
  });

  it("percent: the value at rank max(1, floor(count × value / 100))", () => {
    expect(passing(values, { kind: "top10", top: true, percent: true, value: 40 })).toEqual([40, 50]);
    expect(passing(values, { kind: "top10", top: true, percent: true, value: 1 })).toEqual([50]);
  });

  it("counts filter-hidden and user-hidden rows too", () => {
    const m = list([10, 40, 20], { kind: "top10", top: true, percent: false, value: 1 });
    m.sheets[0] = { ...m.sheets[0]!, rows: m.sheets[0]!.rows.map((r) => (r.index === 2 ? { ...r, hidden: true } : r)) };
    expect(rowPasses(m, m.sheets[0]!, 2, ctx)).toBe(true);
    expect(rowPasses(m, m.sheets[0]!, 1, ctx)).toBe(false);
  });
});

describe("dynamic", () => {
  it("above and below average strictly against the mean of the numeric data cells; non-numeric cells fail", () => {
    expect(passing([10, 20, 30, "x"], { kind: "dynamic", type: "aboveAverage" })).toEqual([30]);
    expect(passing([10, 20, 30, "x"], { kind: "dynamic", type: "belowAverage" })).toEqual([10]);
  });

  it("the date types against today from the clock (Wednesday 2026-10-14), weeks from Sunday", () => {
    setClock(new Date(2026, 9, 14, 15, 30));
    const d = (m: number, day: number): number => date(2026, m, day);
    const dates = [d(10, 14) + 0.75, d(10, 13), d(10, 15), d(10, 11), d(10, 17), d(10, 4), d(10, 24), d(9, 30), d(11, 1), d(7, 1), d(1, 1), date(2025, 10, 14), date(2027, 1, 1)];
    const check = (type: Extract<FilterCriteria, { kind: "dynamic" }>["type"]): CellValue[] => passing(dates, { kind: "dynamic", type }, DATE_STYLE);
    expect(check("today")).toEqual([dates[0]]);
    expect(check("yesterday")).toEqual([d(10, 13)]);
    expect(check("tomorrow")).toEqual([d(10, 15)]);
    expect(check("thisWeek")).toEqual([dates[0], d(10, 13), d(10, 15), d(10, 11), d(10, 17)]);
    expect(check("lastWeek")).toEqual([d(10, 4)]);
    expect(check("nextWeek")).toEqual([d(10, 24)]);
    expect(check("thisMonth")).toEqual([dates[0], d(10, 13), d(10, 15), d(10, 11), d(10, 17), d(10, 4), d(10, 24)]);
    expect(check("lastMonth")).toEqual([d(9, 30)]);
    expect(check("nextMonth")).toEqual([d(11, 1)]);
    expect(check("thisQuarter")).toEqual([dates[0], d(10, 13), d(10, 15), d(10, 11), d(10, 17), d(10, 4), d(10, 24), d(11, 1)]);
    expect(check("lastQuarter")).toEqual([d(9, 30), d(7, 1)]);
    expect(check("nextQuarter")).toEqual([date(2027, 1, 1)]);
    expect(check("thisYear")).toHaveLength(11);
    expect(check("lastYear")).toEqual([date(2025, 10, 14)]);
    expect(check("nextYear")).toEqual([date(2027, 1, 1)]);
    expect(check("yearToDate")).toEqual([dates[0], d(10, 13), d(10, 11), d(10, 4), d(9, 30), d(7, 1), d(1, 1)]);
    expect(check("Q4")).toEqual([dates[0], d(10, 13), d(10, 15), d(10, 11), d(10, 17), d(10, 4), d(10, 24), d(11, 1), date(2025, 10, 14)]);
    expect(check("M1")).toEqual([d(1, 1), date(2027, 1, 1)]);
    expect(passing(["today"], { kind: "dynamic", type: "today" }, DATE_STYLE)).toEqual([]);
  });

  it("reads today in the workbook's date system", () => {
    setClock(new Date(2026, 9, 14));
    const m = list([dateToSerial(2026, 10, 14, "1904")], { kind: "dynamic", type: "today" }, DATE_STYLE);
    m.dateSystem = "1904";
    expect(rowPasses(m, m.sheets[0]!, 1, { ...ctx, dateSystem: "1904" })).toBe(true);
  });
});

describe("applyAutoFilter", () => {
  const twoColumns = (): SpreadsheetModel =>
    model(
      sheet({
        rows: [
          row(0, [cell(0, "Region"), cell(1, "Amount")]),
          row(1, [cell(0, "East"), cell(1, 10)]),
          row(2, [cell(0, "West"), cell(1, 500)]),
          row(3, [cell(0, "East"), cell(1, 900)], { hidden: true }),
          row(4, [cell(0, "South"), cell(1, 600)]),
        ],
        autoFilter: {
          range: range(0, 0, 4, 1),
          columns: [
            { colId: 0, criteria: { kind: "values", values: ["East", "West"], blanks: false, dateGroups: [] } },
            { colId: 1, criteria: { kind: "custom", join: "and", conditions: [{ operator: "greaterThan", value: "100" }] } },
          ],
        },
      }),
    );

  it("sets filterHidden exactly on the rows rowPasses rejects (AND across columns) and never touches hidden", () => {
    const m = twoColumns();
    expect(applyAutoFilter(m, "s1", ctx).changedRows).toEqual([1, 4]);
    expect(hiddenRows(m, "filterHidden")).toEqual([1, 4]);
    expect(hiddenRows(m, "hidden")).toEqual([3]);
    expect(m.sheets[0]!.dirty).toBe(true);
  });

  it("evaluates rows up to the used range only, and a filter-hidden blank row is a Row with no cells", () => {
    const m = model(sheet({ rows: [row(0, [cell(0, "H")]), row(1, [cell(0, "a")]), row(3, [cell(0, "b")])], autoFilter: { range: range(0, 0, 1000, 0), columns: [{ colId: 0, criteria: { kind: "values", values: ["b"], blanks: false, dateGroups: [] } }] } }));
    applyAutoFilter(m, "s1", ctx);
    expect(m.sheets[0]!.rows).toEqual([row(0, [cell(0, "H")]), row(1, [cell(0, "a")], { filterHidden: true }), row(2, [], { filterHidden: true }), row(3, [cell(0, "b")])]);
  });

  it("grows the range's bottom to the header row's current region, never shrinks it", () => {
    const m = list(["a", "b"]);
    m.sheets[0] = { ...m.sheets[0]!, rows: [...m.sheets[0]!.rows, row(3, [cell(0, "c")])], autoFilter: { ...m.sheets[0]!.autoFilter!, sourceXml: "<autoFilter/>" } };
    applyAutoFilter(m, "s1", ctx);
    expect(m.sheets[0]!.autoFilter).toEqual({ range: range(0, 0, 3, 0), columns: [] });
    expect(m.structureDirty).toBe(true);
    const wide = list(["a"]);
    wide.sheets[0] = { ...wide.sheets[0]!, autoFilter: { range: range(0, 0, 50, 0), columns: [] } };
    applyAutoFilter(wide, "s1", ctx);
    expect(wide.sheets[0]!.autoFilter!.range).toEqual(range(0, 0, 50, 0));
  });

  it("keeps the autofilter object when nothing changes", () => {
    const m = list(["a", "b"], { kind: "values", values: ["a", "b"], blanks: false, dateGroups: [] });
    const before = m.sheets[0];
    expect(applyAutoFilter(m, "s1", ctx).changedRows).toEqual([]);
    expect(m.sheets[0]).toBe(before);
  });

  it("a preserved criterion keeps each row's existing filterHidden and adds the modelled rejections", () => {
    const m = model(
      sheet({
        rows: [row(0, [cell(0, "Item"), cell(1, "Qty")]), row(1, [cell(0, "a"), cell(1, 1)], { filterHidden: true }), row(2, [cell(0, "b"), cell(1, 2)]), row(3, [cell(0, "c"), cell(1, 3)])],
        autoFilter: {
          range: range(0, 0, 3, 1),
          columns: [
            { colId: 0, criteria: { kind: "preserved", xml: "<colorFilter dxfId=\"0\"/>", label: "colour" } },
            { colId: 1, criteria: { kind: "custom", join: "and", conditions: [{ operator: "lessThan", value: "3" }] } },
          ],
        },
      }),
    );
    expect(reapplyFilter(m, "s1", ctx)).toEqual({ ok: true, changedRows: [3] });
    expect(hiddenRows(m, "filterHidden")).toEqual([1, 3]);
  });
});

describe("the read rule — classifyHiddenRows", () => {
  const read = (filter: AutoFilter, rows: Row[]): Row[] => {
    const s = sheet({ rows, autoFilter: filter });
    return classifyHiddenRows(model(s), s, ctx);
  };
  const regions = (): Row[] => [
    row(0, [cell(0, "Region")]),
    row(1, [cell(0, "East")]),
    row(2, [cell(0, "South")], { hidden: true }),
    row(3, [cell(0, "East")], { hidden: true }),
    row(4, [cell(0, "South")]),
    row(7, [cell(0, "South")], { hidden: true }),
  ];
  const east: AutoFilter = { range: range(0, 0, 4, 0), columns: [{ colId: 0, criteria: { kind: "values", values: ["East"], blanks: false, dateGroups: [] } }] };

  it("a hidden data row a modelled criterion rejects is filter-hidden, one it passes user-hidden; visible rows stay visible; rows outside stay hidden", () => {
    const rows = read(east, regions());
    expect(rows.find((r) => r.index === 2)).toEqual(row(2, [cell(0, "South")], { filterHidden: true }));
    expect(rows.find((r) => r.index === 3)).toEqual(row(3, [cell(0, "East")], { hidden: true }));
    expect(rows.find((r) => r.index === 4)).toEqual(row(4, [cell(0, "South")]));
    expect(rows.find((r) => r.index === 7)).toEqual(row(7, [cell(0, "South")], { hidden: true }));
  });

  it("with a preserved criterion every hidden data row is filter-hidden", () => {
    const rows = read({ range: range(0, 0, 4, 0), columns: [{ colId: 0, criteria: { kind: "preserved", xml: "<iconFilter/>", label: "icon" } }] }, regions());
    expect(rows.filter((r) => r.filterHidden).map((r) => r.index)).toEqual([2, 3]);
    expect(rows.filter((r) => r.hidden).map((r) => r.index)).toEqual([7]);
  });

  it("a sheet without an autofilter keeps its rows", () => {
    const s = sheet({ rows: regions() });
    expect(classifyHiddenRows(model(s), s, ctx)).toBe(s.rows);
  });
});

describe("the commands on a sheet holding both kinds of hidden row", () => {
  const both = (): SpreadsheetModel =>
    model(
      sheet({
        rows: [
          row(0, [cell(0, "H")]),
          row(1, [cell(0, "a")], { filterHidden: true }),
          row(2, [cell(0, "b")], { hidden: true }),
          row(3, [cell(0, "c")], { hidden: true, filterHidden: true }),
          row(4, [], { filterHidden: true }),
          row(5, [cell(0, "d")]),
        ],
        autoFilter: { range: range(0, 0, 5, 0), columns: [{ colId: 0, criteria: { kind: "values", values: ["b", "d"], blanks: false, dateGroups: [] }, hideButton: true }], sourceXml: "<autoFilter/>" },
      }),
    );

  it("Clear filters removes every criterion, keeps the buttons, clears filterHidden and never touches hidden", () => {
    const m = both();
    expect(clearFilters(m, "s1")).toEqual({ ok: true, changedRows: [1, 3, 4] });
    expect(m.sheets[0]!.autoFilter).toEqual({ range: range(0, 0, 5, 0), columns: [{ colId: 0, hideButton: true }] });
    expect(hiddenRows(m, "filterHidden")).toEqual([]);
    expect(hiddenRows(m, "hidden")).toEqual([2, 3]);
    expect(findRow(m.sheets[0]!, 4)).toBeUndefined();
  });

  it("turning the filter off removes autoFilter and every filterHidden, never hidden", () => {
    const m = both();
    expect(toggleAutoFilter(m, "s1", null, ctx)).toEqual({ ok: true, changedRows: [1, 3, 4] });
    expect(m.sheets[0]!.autoFilter).toBeUndefined();
    expect(hiddenRows(m, "filterHidden")).toEqual([]);
    expect(hiddenRows(m, "hidden")).toEqual([2, 3]);
    expect(m.structureDirty).toBe(true);
  });

  it("turning it on covers the range given, with no criteria; an empty region refuses", () => {
    const m = list(["a"]);
    m.sheets[0] = { ...m.sheets[0]!, autoFilter: undefined };
    expect(toggleAutoFilter(m, "s1", null, ctx)).toEqual({ ok: false, message: EMPTY_FILTER_REGION_MESSAGE });
    expect(toggleAutoFilter(m, "s1", range(0, 0, 1, 0), ctx)).toEqual({ ok: true, changedRows: [] });
    expect(m.sheets[0]!.autoFilter).toEqual({ range: range(0, 0, 1, 0), columns: [] });
    expect(EMPTY_FILTER_REGION_MESSAGE).toBe("Select a cell inside the data you want to filter");
  });

  it("setFilterCriteria replaces the column's criteria (keeping its button flag) and applies", () => {
    const m = both();
    const result = setFilterCriteria(m, "s1", 0, { kind: "values", values: ["a"], blanks: false, dateGroups: [] }, ctx);
    expect(result.ok).toBe(true);
    expect(m.sheets[0]!.autoFilter!.columns).toEqual([{ colId: 0, criteria: { kind: "values", values: ["a"], blanks: false, dateGroups: [] }, hideButton: true }]);
    expect(m.sheets[0]!.autoFilter!.sourceXml).toBeUndefined();
    expect(hiddenRows(m, "filterHidden")).toEqual([2, 3, 4, 5]);
    expect(hiddenRows(m, "hidden")).toEqual([2, 3]);
    setFilterCriteria(m, "s1", 0, undefined, ctx);
    expect(m.sheets[0]!.autoFilter!.columns).toEqual([{ colId: 0, hideButton: true }]);
    expect(hiddenRows(m, "filterHidden")).toEqual([]);
  });

  it("an ods filter that does not decompose is replaced by any column's criteria", () => {
    const m = both();
    m.sheets[0] = { ...m.sheets[0]!, autoFilter: { range: range(0, 0, 5, 0), columns: [], odsUnmappedFilter: "<table:filter/>" } };
    expect(hasCriteria(m.sheets[0]!.autoFilter)).toBe(true);
    setFilterCriteria(m, "s1", 0, undefined, ctx);
    expect(m.sheets[0]!.autoFilter).toEqual({ range: range(0, 0, 5, 0), columns: [] });
    expect(hiddenRows(m, "filterHidden")).toEqual([]);
  });

  it("refuses on a protected sheet unless autoFilter=\"0\"", () => {
    const m = both();
    m.sheets[0] = { ...m.sheets[0]!, protection: { sheet: true } };
    expect(clearFilters(m, "s1")).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    m.sheets[0] = { ...m.sheets[0]!, protection: { sheet: true, allowAutoFilter: true } };
    expect(clearFilters(m, "s1").ok).toBe(true);
  });
});

describe("visibleRowSpans", () => {
  it("splits a range around its filter-hidden rows only", () => {
    const s = sheet({ rows: [row(2, [], { filterHidden: true }), row(3, [], { filterHidden: true }), row(5, [], { hidden: true }), row(7, [], { filterHidden: true })] });
    expect(visibleRowSpans(s, range(0, 0, 9, 1))).toEqual([range(0, 0, 1, 1), range(4, 0, 6, 1), range(8, 0, 9, 1)]);
    expect(visibleRowSpans(s, range(2, 0, 3, 1))).toEqual([]);
    expect(visibleRowSpans(sheet(), range(0, 0, 3, 3))).toEqual([range(0, 0, 3, 3)]);
  });
});

describe("filterValueList", () => {
  const regions = (): SpreadsheetModel =>
    model(
      sheet({
        rows: [
          row(0, [cell(0, "Region"), cell(1, "Amount")]),
          row(1, [cell(0, "West"), cell(1, 5)]),
          row(2, [cell(0, "east"), cell(1, 50)]),
          row(3, [cell(1, 500)]),
          row(4, [cell(0, 10), cell(1, 5)]),
          row(5, [cell(0, "East"), cell(1, 5)]),
          row(6, [cell(0, "North"), cell(1, 5000)]),
        ],
        autoFilter: {
          range: range(0, 0, 6, 1),
          columns: [
            { colId: 0, criteria: { kind: "values", values: ["East"], blanks: true, dateGroups: [] } },
            { colId: 1, criteria: { kind: "custom", join: "and", conditions: [{ operator: "lessThan", value: "1000" }] } },
          ],
        },
      }),
    );

  it("one item per distinct display text of the rows the other columns pass, in the sort order, blanks apart", () => {
    const m = regions();
    expect(filterValueList(m, m.sheets[0]!, 0, ctx)).toEqual({
      items: [
        { text: "10", checked: false },
        { text: "east", checked: true },
        { text: "West", checked: false },
      ],
      blanks: true,
      truncated: false,
    });
  });

  it("with no criteria every item is checked; blanks null when no row is blank", () => {
    const m = regions();
    expect(filterValueList(m, m.sheets[0]!, 1, ctx)).toEqual({
      items: [
        { text: "5", checked: true },
        { text: "50", checked: true },
        { text: "500", checked: true },
      ],
      blanks: null,
      truncated: false,
    });
  });

  it(`stops at FILTER_VALUE_LIST_MAX = ${FILTER_VALUE_LIST_MAX}`, () => {
    const values = Array.from({ length: FILTER_VALUE_LIST_MAX + 5 }, (_, i) => i);
    const m = list(values);
    const out = filterValueList(m, m.sheets[0]!, 0, ctx);
    expect(FILTER_VALUE_LIST_MAX).toBe(10_000);
    expect(out.items).toHaveLength(FILTER_VALUE_LIST_MAX);
    expect(out.truncated).toBe(true);
  });
});

/** A selection of `ranges`, active at the first one's top-left. */
function at(ranges: GridRange[]): SheetSelection {
  return { active: { row: ranges[0]!.top, col: ranges[0]!.left }, ranges, activeRangeIndex: 0 };
}

describe("choosing the range", () => {
  const s = sheet({
    rows: [
      row(1, [cell(1, "Region"), cell(2, "Amount")]),
      row(2, [cell(1, "East"), cell(2, 1)]),
      row(3, [cell(1, "West"), cell(2, 2)]),
    ],
  });
  const nav: NavContext = { sheet: s, merges: buildMergeIndex([]), isRowHidden: () => false, isColHidden: () => false, pageRows: 10, pageCols: 5 };

  it("a single cell → its current region; an empty one refuses", () => {
    expect(autoFilterRangeFor(at([range(2, 2, 2, 2)]), nav)).toEqual(range(1, 1, 3, 2));
    expect(autoFilterRangeFor(at([range(9, 9, 9, 9)]), nav)).toBeNull();
  });

  it("a one-row selection → its columns down to the region's bottom; a multi-row selection → itself", () => {
    expect(autoFilterRangeFor(at([range(1, 1, 1, 1)]), nav)).toEqual(range(1, 1, 3, 2));
    expect(autoFilterRangeFor(at([range(1, 1, 1, 4)]), nav)).toEqual(range(1, 1, 3, 4));
    expect(autoFilterRangeFor(at([range(1, 1, 2, 1)]), nav)).toEqual(range(1, 1, 2, 1));
  });
});

describe("what the UI reads", () => {
  it("filterRecordCounts: N of M while a criterion exists; M stops at the used range", () => {
    const m = list(["a", "b", "c"], { kind: "values", values: ["a"], blanks: false, dateGroups: [] });
    m.sheets[0] = { ...m.sheets[0]!, autoFilter: { ...m.sheets[0]!.autoFilter!, range: range(0, 0, 999, 0) } };
    expect(filterRecordCounts(m.sheets[0]!)).toEqual({ found: 3, total: 3 });
    applyAutoFilter(m, "s1", ctx);
    expect(filterRecordCounts(m.sheets[0]!)).toEqual({ found: 1, total: 3 });
    expect(filterRecordCounts(list(["a"]).sheets[0]!)).toBeNull();
  });

  it("filterButtonState: sorted beats filtered", () => {
    const filter: AutoFilter = {
      range: range(0, 2, 9, 4),
      columns: [{ colId: 1, criteria: { kind: "values", values: [], blanks: true, dateGroups: [] } }],
      sort: { ref: range(1, 2, 9, 4), keys: [{ col: 4, descending: true }] },
    };
    expect(filterButtonState(filter, 0)).toBe("none");
    expect(filterButtonState(filter, 1)).toBe("filtered");
    expect(filterButtonState(filter, 2)).toBe("sorted-desc");
    expect(filterButtonState({ ...filter, sort: { ref: range(1, 2, 9, 4), keys: [{ col: 3, descending: false }] } }, 1)).toBe("sorted-asc");
  });

  it("columnValueKind: the majority of the non-blank data cells, ties → text", () => {
    const s = sheet({
      rows: [
        row(1, [cell(0, 1), cell(1, date(2026, 1, 1), { styleId: DATE_STYLE }), cell(2, "a")]),
        row(2, [cell(0, 2), cell(1, date(2026, 1, 2), { styleId: DATE_STYLE }), cell(2, 1)]),
        row(3, [cell(0, "x"), cell(1, 3)]),
      ],
    });
    const m = model(s);
    expect(columnValueKind(m, s, 0, { top: 1, bottom: 3 }, ctx)).toBe("number");
    expect(columnValueKind(m, s, 1, { top: 1, bottom: 3 }, ctx)).toBe("date");
    expect(columnValueKind(m, s, 2, { top: 1, bottom: 3 }, ctx)).toBe("text");
  });
});
