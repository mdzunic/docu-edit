// The autofilter (SPEC-128 "The autofilter model", "Criteria semantics",
// "Filter-hidden or user-hidden", "Choosing the range") — THIS FILE is the
// normative source for the model, each criterion's matching rule and the
// filter commands.
import type { GridRange } from "./address";
import type { Row, Sheet, SpreadsheetModel } from "./model";
import type { FormatContext } from "./numberFormat/render";
import type { NavContext, SheetSelection } from "./selection";
import type { SortState } from "./sort";

export interface AutoFilter {
  range: GridRange; // header row = range.top; data rows = top+1..bottom
  columns: FilterColumn[]; // ascending colId, at most one per colId
  sort?: SortState;
  /** xlsx `<autoFilter>` / ods `<table:database-range>`, verbatim while unchanged. */
  sourceXml?: string;
  /** ods: the database range's name ("__Anonymous_Sheet_DB__0"). */
  odsName?: string;
  /** ods: a `<table:filter>` that does not decompose per column — kept
   *  verbatim; its columns offer only "Clear filter". */
  odsUnmappedFilter?: string;
}

export interface FilterColumn {
  colId: number; // offset from range.left
  criteria?: FilterCriteria;
  hideButton?: true;
}

export type FilterCriteria =
  | { kind: "values"; values: string[]; blanks: boolean; dateGroups: DateGroupItem[] }
  | { kind: "custom"; join: "and" | "or"; conditions: [FilterCondition] | [FilterCondition, FilterCondition] }
  | { kind: "top10"; top: boolean; percent: boolean; value: number }
  | { kind: "dynamic"; type: DynamicFilterType }
  | { kind: "preserved"; xml: string; label: "colour" | "icon" };

export type FilterOperator = "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual";

export interface FilterCondition {
  operator: FilterOperator;
  value: string; // wire text: invariant number text, or text with wildcards * ? ~
}

export interface DateGroupItem {
  grouping: "year" | "month" | "day" | "hour" | "minute" | "second";
  year: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
}

export type DynamicFilterType =
  | "aboveAverage"
  | "belowAverage"
  | "today"
  | "yesterday"
  | "tomorrow"
  | "thisWeek"
  | "lastWeek"
  | "nextWeek"
  | "thisMonth"
  | "lastMonth"
  | "nextMonth"
  | "thisQuarter"
  | "lastQuarter"
  | "nextQuarter"
  | "thisYear"
  | "lastYear"
  | "nextYear"
  | "yearToDate"
  | "Q1"
  | "Q2"
  | "Q3"
  | "Q4"
  | "M1"
  | "M2"
  | "M3"
  | "M4"
  | "M5"
  | "M6"
  | "M7"
  | "M8"
  | "M9"
  | "M10"
  | "M11"
  | "M12";

export const FILTER_VALUE_LIST_MAX = 10_000;

/** What a filter command did: the rows whose `filterHidden` changed, or why
 * it changed nothing. */
export type FilterCommandResult = { ok: true; changedRows: number[] } | { ok: false; message: string };

export const EMPTY_FILTER_REGION_MESSAGE = "Select a cell inside the data you want to filter";

/** Whether a row passes every modelled criterion of the sheet's autofilter
 * (AND across columns; a preserved criterion passes). */
export function rowPasses(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, row: number, ctx: FormatContext): boolean {
  void model;
  void sheet;
  void row;
  void ctx;
  throw new Error("not implemented");
}

/** Sets `filterHidden` on the autofilter's data rows per its criteria. */
export function applyAutoFilter(model: SpreadsheetModel, sheetId: string, ctx: FormatContext): { changedRows: number[] } {
  void model;
  void sheetId;
  void ctx;
  throw new Error("not implemented");
}

export function filterValueList(
  model: Pick<SpreadsheetModel, "styles">,
  sheet: Sheet,
  colId: number,
  ctx: FormatContext,
): { items: { text: string; checked: boolean }[]; blanks: boolean | null; truncated: boolean } {
  void model;
  void sheet;
  void colId;
  void ctx;
  throw new Error("not implemented");
}

/** The parts of `range` on rows not filter-hidden, as row spans. */
export function visibleRowSpans(sheet: Sheet, range: GridRange): GridRange[] {
  void sheet;
  void range;
  throw new Error("not implemented");
}

/** The read rule ("Filter-hidden or user-hidden"): the sheet's rows with
 * every `hidden` row inside its autofilter's data rows reclassified as
 * `filterHidden` when a modelled criterion rejects it or the filter carries
 * a preserved criterion. The same array when nothing changes. */
export function classifyHiddenRows(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, ctx: FormatContext): Row[] {
  void model;
  void sheet;
  void ctx;
  throw new Error("not implemented");
}

/** "Choosing the range": the range Data ▸ Filter covers, or null. */
export function autoFilterRangeFor(selection: SheetSelection, nav: NavContext): GridRange | null {
  void selection;
  void nav;
  throw new Error("not implemented");
}
