// The autofilter (SPEC-128 "The autofilter model", "Criteria semantics",
// "Filter-hidden or user-hidden", "Choosing the range") — THIS FILE is the
// normative source for the model, each criterion's matching rule, the
// distinction between filter-hidden and user-hidden rows and the filter
// commands.
//
// - The model. `Sheet.autoFilter` covers `range` — its header row is
//   `range.top`, its data rows `top + 1 .. bottom` — with at most one
//   FilterColumn per `colId` (an offset from `range.left`), and the record of
//   its last sort. One autofilter per sheet. A row hidden by the filter
//   carries `Row.filterHidden`, distinct from the user's `Row.hidden`: the
//   row shows hidden when either is set (model.ts rowIsHidden), Unhide clears
//   only `hidden`, and clearing or removing the filter clears only
//   `filterHidden`. A filter-hidden blank row is a Row with no cells.
// - Display text. A cell's display text is plainText(formatValue(value, its
//   number format, ctx)) with no `widthPx` (never "####"), compared
//   case-insensitively; a blank is a cell whose display text is "". A date
//   cell is a number whose format's first section has kind "date"; a
//   column's type (labels) is the majority of its non-blank data cells —
//   date cells, other numbers, or everything else (ties → text).
// - Criteria. `values` passes a display text equal to an entry, a blank when
//   `blanks`, a date cell whose serial's parts equal a DateGroupItem down to
//   its grouping; nothing listed hides every data row. `custom` — per
//   condition: `notEqual " "` is "non-blank" (Excel's encoding); an
//   unescaped `*` or `?` under equal/notEqual is a wildcard match on the
//   display text (`~` escapes); else a value that parses as an invariant
//   number against a numeric cell compares numerically; else the display
//   text compares through sortCollator(tag, false) — a numeric cell never
//   matches a text ordering. A blank fails every operator but notEqual;
//   `join` combines the two. `top10` ranks the numeric cells of every data
//   row (visible or not): the threshold is the `value`-th largest (smallest)
//   — for `percent` the one at rank max(1, floor(count × value / 100)) — and
//   a cell passes at or beyond it. `dynamic`: above/below average strictly
//   against the mean of the numeric data cells; the date types compare the
//   serial's date part with today (core/clock.ts now(), local time, through
//   dateToSerial in the workbook's date system) — weeks start on Sunday,
//   calendar quarters, yearToDate is 1 January ≤ d ≤ today, Qn/Mn match any
//   year; non-numeric cells fail. `preserved` (colour and icon filters) and
//   an ods `odsUnmappedFilter` are not evaluated: applying keeps each row's
//   existing `filterHidden` and adds the modelled criteria's rejections.
// - Applying. The range's bottom first grows to its header row's
//   currentRegion bottom when that is lower (never shrinks); rows from
//   `top + 1` to min(bottom, usedRange.bottom) are evaluated — trailing rows
//   past the used range are never hidden — and `filterHidden` is set exactly
//   on the rejected ones (AND across columns), cleared on every other row.
// - Reading never changes what is visible (classifyHiddenRows): a hidden row
//   inside the data rows is `filterHidden` when a modelled criterion rejects
//   it or the filter carries a preserved criterion, else `hidden`.
// - Commands replace — never mutate — `autoFilter` (dropping its
//   `sourceXml`), set `Sheet.dirty`, and set `structureDirty` when the
//   autofilter is created, removed or its range moves (the xlsx workbook part
//   carries `_xlnm._FilterDatabase`).
import { findRow, getCell, usedRange, type CellRef, type GridRange } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import { now } from "../core/clock";
import { buildMergeIndex } from "./merges";
import { type Cell, type CellValue, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { parseFormatCode } from "./numberFormat/parse";
import { formatValue, plainText, type FormatContext } from "./numberFormat/render";
import { currentRegion, type NavContext, type SheetSelection } from "./selection";
import { dateToSerial, serialToParts } from "./serial";
import { withRows } from "./sheetEdit";
import { compareSortValues, sortCollator, type SortState } from "./sort";
import { cellStyle, effectiveStyleId } from "./styles";

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

// --- display text --------------------------------------------------------------

/** Whether a format code's first section renders a date. */
const dateCodes = new Map<string, boolean>();
function isDateCode(code: string | undefined): boolean {
  if (code === undefined) {
    return false;
  }
  let date = dateCodes.get(code);
  if (date === undefined) {
    try {
      date = parseFormatCode(code).sections[0]?.kind === "date";
    } catch {
      date = false;
    }
    dateCodes.set(code, date);
  }
  return date;
}

/** One cell as the criteria read it. */
interface CellView {
  value: CellValue;
  text: string; // display text, lower-cased
  date: boolean;
}

const BLANK_VIEW: CellView = { value: null, text: "", date: false };

function cellView(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, row: number, col: number, ctx: FormatContext): CellView {
  const value = getCell(sheet, row, col)?.value ?? null;
  if (value === null) {
    return BLANK_VIEW;
  }
  const code = formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row, col)), ctx.locale);
  const shown = plainText(formatValue(value, code, { ...ctx, widthPx: undefined }));
  return { value, text: shown.toLowerCase(), date: typeof value === "number" && isDateCode(code) };
}

/** The display text of a cell as shown (not lower-cased). */
function displayText(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, row: number, col: number, ctx: FormatContext): string {
  const value = getCell(sheet, row, col)?.value ?? null;
  if (value === null) {
    return "";
  }
  const code = formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row, col)), ctx.locale);
  return plainText(formatValue(value, code, { ...ctx, widthPx: undefined }));
}

/** The last data row evaluated: min(range.bottom, usedRange.bottom). */
function lastDataRow(sheet: Sheet, filter: AutoFilter): number {
  return Math.min(filter.range.bottom, usedRange(sheet)?.bottom ?? filter.range.top);
}

// --- criteria ---------------------------------------------------------------------

/** A wildcard pattern (`*`, `?`, `~` escapes) as a whole-text RegExp, or
 * null when it holds no unescaped wildcard. */
function wildcardPattern(value: string): RegExp | null {
  let source = "";
  let wild = false;
  for (let i = 0; i < value.length; i += 1) {
    const ch = value[i]!;
    if (ch === "~" && i + 1 < value.length) {
      source += value[i + 1]!.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      i += 1;
    } else if (ch === "*") {
      source += "[\\s\\S]*";
      wild = true;
    } else if (ch === "?") {
      source += "[\\s\\S]";
      wild = true;
    } else {
      source += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
  }
  return wild ? new RegExp(`^${source}$`, "i") : null;
}

/** `~x` escapes removed (a value with no wildcard compares literally). */
function unescapeWildcards(value: string): string {
  return value.replace(/~([\s\S])/g, "$1");
}

/** An invariant number text as a number, or null. */
function invariantNumber(text: string): number | null {
  const trimmed = text.trim();
  if (trimmed === "" || !/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(trimmed)) {
    return null;
  }
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : null;
}

function compareOp(order: number, operator: FilterOperator): boolean {
  switch (operator) {
    case "equal":
      return order === 0;
    case "notEqual":
      return order !== 0;
    case "greaterThan":
      return order > 0;
    case "greaterThanOrEqual":
      return order >= 0;
    case "lessThan":
      return order < 0;
    case "lessThanOrEqual":
      return order <= 0;
  }
}

function conditionPasses(cell: CellView, condition: FilterCondition, collator: Intl.Collator): boolean {
  const { operator, value } = condition;
  if (operator === "notEqual" && value === " ") {
    return cell.text !== "";
  }
  if (cell.text === "") {
    return operator === "notEqual";
  }
  if (operator === "equal" || operator === "notEqual") {
    const pattern = wildcardPattern(value);
    if (pattern) {
      return pattern.test(cell.text) === (operator === "equal");
    }
  }
  const number = invariantNumber(value);
  if (number !== null && typeof cell.value === "number") {
    return compareOp(cell.value < number ? -1 : cell.value > number ? 1 : 0, operator);
  }
  if (typeof cell.value === "number" && operator !== "equal" && operator !== "notEqual") {
    return false;
  }
  return compareOp(collator.compare(cell.text, unescapeWildcards(value).toLowerCase()), operator);
}

function dateGroupMatches(serial: number, item: DateGroupItem, dateSystem: "1900" | "1904"): boolean {
  const parts = serialToParts(serial, dateSystem);
  const order = ["year", "month", "day", "hour", "minute", "second"] as const;
  for (const part of order) {
    if (parts[part] !== item[part]) {
      return false;
    }
    if (part === item.grouping) {
      return true;
    }
  }
  return true;
}

/** Today's serial (date part) and its calendar parts, in `system`. */
function today(system: "1900" | "1904"): { serial: number; year: number; month: number; day: number; weekday: number } {
  const at = now();
  const year = at.getFullYear();
  const month = at.getMonth() + 1;
  const day = at.getDate();
  return { serial: dateToSerial(year, month, day, system), year, month, day, weekday: at.getDay() };
}

const monthIndex = (year: number, month: number): number => year * 12 + month - 1;
const quarterIndex = (year: number, month: number): number => year * 4 + Math.floor((month - 1) / 3);

function dynamicPasses(cell: CellView, type: DynamicFilterType, stats: ColumnStats, system: "1900" | "1904"): boolean {
  if (typeof cell.value !== "number") {
    return false;
  }
  const value = cell.value;
  if (type === "aboveAverage" || type === "belowAverage") {
    const mean = stats.mean();
    return mean !== null && (type === "aboveAverage" ? value > mean : value < mean);
  }
  const d = Math.floor(value);
  const parts = serialToParts(d, system);
  const t = today(system);
  const weekStart = t.serial - t.weekday;
  switch (type) {
    case "today":
      return d === t.serial;
    case "yesterday":
      return d === t.serial - 1;
    case "tomorrow":
      return d === t.serial + 1;
    case "thisWeek":
      return d >= weekStart && d < weekStart + 7;
    case "lastWeek":
      return d >= weekStart - 7 && d < weekStart;
    case "nextWeek":
      return d >= weekStart + 7 && d < weekStart + 14;
    case "thisMonth":
      return monthIndex(parts.year, parts.month) === monthIndex(t.year, t.month);
    case "lastMonth":
      return monthIndex(parts.year, parts.month) === monthIndex(t.year, t.month) - 1;
    case "nextMonth":
      return monthIndex(parts.year, parts.month) === monthIndex(t.year, t.month) + 1;
    case "thisQuarter":
      return quarterIndex(parts.year, parts.month) === quarterIndex(t.year, t.month);
    case "lastQuarter":
      return quarterIndex(parts.year, parts.month) === quarterIndex(t.year, t.month) - 1;
    case "nextQuarter":
      return quarterIndex(parts.year, parts.month) === quarterIndex(t.year, t.month) + 1;
    case "thisYear":
      return parts.year === t.year;
    case "lastYear":
      return parts.year === t.year - 1;
    case "nextYear":
      return parts.year === t.year + 1;
    case "yearToDate":
      return d >= dateToSerial(t.year, 1, 1, system) && d <= t.serial;
    default: {
      const quarter = /^Q([1-4])$/.exec(type);
      if (quarter) {
        return Math.floor((parts.month - 1) / 3) + 1 === Number(quarter[1]);
      }
      const month = /^M(\d{1,2})$/.exec(type);
      return month !== null && parts.month === Number(month[1]);
    }
  }
}

/** A column's numeric data cells, ranked and averaged on first use. */
class ColumnStats {
  private numbers: number[] | null = null;

  constructor(
    private readonly sheet: Sheet,
    private readonly col: number,
    private readonly top: number,
    private readonly bottom: number,
  ) {}

  private sorted(): number[] {
    if (this.numbers === null) {
      const out: number[] = [];
      for (const row of this.sheet.rows) {
        if (row.index < this.top || row.index > this.bottom) {
          continue;
        }
        const value = row.cells.find((cell) => cell.col === this.col)?.value;
        if (typeof value === "number") {
          out.push(value);
        }
      }
      this.numbers = out.toSorted((a, b) => a - b);
    }
    return this.numbers;
  }

  mean(): number | null {
    const numbers = this.sorted();
    return numbers.length === 0 ? null : numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
  }

  /** The top10 threshold, or null when there is no numeric cell. */
  threshold(criteria: Extract<FilterCriteria, { kind: "top10" }>): number | null {
    const numbers = this.sorted();
    if (numbers.length === 0) {
      return null;
    }
    const rank = criteria.percent ? Math.max(1, Math.floor((numbers.length * criteria.value) / 100)) : Math.max(1, Math.floor(criteria.value));
    const at = Math.min(rank, numbers.length) - 1;
    return criteria.top ? numbers[numbers.length - 1 - at]! : numbers[at]!;
  }
}

/** Evaluates one autofilter's modelled criteria row by row. */
class FilterEvaluator {
  readonly filter: AutoFilter;
  readonly hasPreserved: boolean;
  private readonly collator: Intl.Collator;
  private readonly stats = new Map<number, ColumnStats>();
  private readonly values = new Map<number, { texts: Set<string> }>();

  constructor(
    private readonly model: Pick<SpreadsheetModel, "styles">,
    private readonly sheet: Sheet,
    filter: AutoFilter,
    private readonly ctx: FormatContext,
  ) {
    this.filter = filter;
    this.hasPreserved = filter.odsUnmappedFilter !== undefined || filter.columns.some((column) => column.criteria?.kind === "preserved");
    this.collator = sortCollator(ctx.locale.tag, false);
  }

  private statsFor(colId: number): ColumnStats {
    let stats = this.stats.get(colId);
    if (!stats) {
      stats = new ColumnStats(this.sheet, this.filter.range.left + colId, this.filter.range.top + 1, lastDataRow(this.sheet, this.filter));
      this.stats.set(colId, stats);
    }
    return stats;
  }

  /** Whether one column's criteria pass a row (preserved: true). */
  columnPasses(colId: number, criteria: FilterCriteria, row: number): boolean {
    const cell = cellView(this.model, this.sheet, row, this.filter.range.left + colId, this.ctx);
    return this.criteriaPass(colId, criteria, cell);
  }

  criteriaPass(colId: number, criteria: FilterCriteria, cell: CellView): boolean {
    switch (criteria.kind) {
      case "values": {
        let entry = this.values.get(colId);
        if (!entry) {
          entry = { texts: new Set(criteria.values.map((value) => value.toLowerCase())) };
          this.values.set(colId, entry);
        }
        if (cell.text === "") {
          return criteria.blanks;
        }
        if (entry.texts.has(cell.text)) {
          return true;
        }
        return cell.date && criteria.dateGroups.some((item) => dateGroupMatches(cell.value as number, item, this.ctx.dateSystem));
      }
      case "custom": {
        const [first, second] = criteria.conditions;
        const a = conditionPasses(cell, first, this.collator);
        if (!second) {
          return a;
        }
        const b = conditionPasses(cell, second, this.collator);
        return criteria.join === "and" ? a && b : a || b;
      }
      case "top10": {
        if (typeof cell.value !== "number") {
          return false;
        }
        const threshold = this.statsFor(colId).threshold(criteria);
        return threshold !== null && (criteria.top ? cell.value >= threshold : cell.value <= threshold);
      }
      case "dynamic":
        return dynamicPasses(cell, criteria.type, this.statsFor(colId), this.ctx.dateSystem);
      case "preserved":
        return true;
    }
  }

  /** Whether a row passes every modelled criterion. */
  passes(row: number): boolean {
    for (const column of this.filter.columns) {
      if (column.criteria && !this.columnPasses(column.colId, column.criteria, row)) {
        return false;
      }
    }
    return true;
  }
}

/** Whether a row passes every modelled criterion of the sheet's autofilter
 * (AND across columns; a preserved criterion passes). */
export function rowPasses(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, row: number, ctx: FormatContext): boolean {
  const filter = sheet.autoFilter;
  return !filter || new FilterEvaluator(model, sheet, filter, ctx).passes(row);
}

/** Whether a filter holds any criterion (the status bar's and the row
 * headers' "filtered" state). */
export function hasCriteria(filter: AutoFilter | undefined): boolean {
  return filter !== undefined && (filter.odsUnmappedFilter !== undefined || filter.columns.some((column) => column.criteria !== undefined));
}

// --- applying -------------------------------------------------------------------

/** A row with `filterHidden` set or cleared (undefined: the row stores nothing). */
function withFilterHidden(row: Row | undefined, index: number, hidden: boolean): Row | undefined {
  if (hidden) {
    return row ? { ...row, filterHidden: true } : { index, cells: [], filterHidden: true };
  }
  if (!row) {
    return undefined;
  }
  const { filterHidden: _gone, ...rest } = row;
  return rest.cells.length === 0 && Object.keys(rest).every((key) => key === "index" || key === "cells" || key === "sourceStyleName") ? undefined : rest;
}

/** Replaces sheet `sheetId` with `next(sheet)`. */
function replaceSheet(model: SpreadsheetModel, sheetId: string, next: (sheet: Sheet) => Sheet): void {
  model.sheets = model.sheets.map((sheet) => (sheet.id === sheetId ? next(sheet) : sheet));
}

/** A NavContext for growing a region (currentRegion reads rows and merges). */
function regionNav(sheet: Sheet): NavContext {
  return { sheet, merges: buildMergeIndex(sheet.merges), isRowHidden: () => false, isColHidden: () => false, pageRows: 1, pageCols: 1 };
}

/** The rows of `sheet` with `filterHidden` cleared everywhere: the same
 * array when none is set, plus the changed indices. */
function clearedRows(sheet: Sheet): { rows: Row[]; changed: number[] } {
  const replaced = new Map<number, Row | undefined>();
  for (const row of sheet.rows) {
    if (row.filterHidden) {
      replaced.set(row.index, withFilterHidden(row, row.index, false));
    }
  }
  return { rows: withRows(sheet.rows, replaced), changed: [...replaced.keys()] };
}

/** Sets `filterHidden` on the autofilter's data rows per its criteria (see
 * the module header). */
export function applyAutoFilter(model: SpreadsheetModel, sheetId: string, ctx: FormatContext): { changedRows: number[] } {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  let filter = sheet?.autoFilter;
  if (!sheet || !filter) {
    return { changedRows: [] };
  }
  const region = currentRegion({ row: filter.range.top, col: filter.range.left }, regionNav(sheet));
  if (region.bottom > filter.range.bottom) {
    const { sourceXml: _source, ...rest } = filter;
    filter = { ...rest, range: { ...filter.range, bottom: region.bottom } };
    model.structureDirty = true;
  }
  const evaluator = new FilterEvaluator(model, sheet, filter, ctx);
  const first = filter.range.top + 1;
  const last = lastDataRow(sheet, filter);
  const replaced = new Map<number, Row | undefined>();
  for (let r = first; r <= last; r += 1) {
    const row = findRow(sheet, r);
    const hidden = !evaluator.passes(r) || (evaluator.hasPreserved && row?.filterHidden === true);
    if (hidden !== (row?.filterHidden === true)) {
      replaced.set(r, withFilterHidden(row, r, hidden));
    }
  }
  // A filter-hidden row anywhere else (past the used range, outside the data rows) shows.
  for (const row of sheet.rows) {
    if (row.filterHidden && (row.index < first || row.index > last)) {
      replaced.set(row.index, withFilterHidden(row, row.index, false));
    }
  }
  const finalFilter = filter;
  if (replaced.size > 0 || finalFilter !== sheet.autoFilter) {
    replaceSheet(model, sheetId, (each) => ({ ...each, rows: withRows(each.rows, replaced), autoFilter: finalFilter, dirty: true }));
  }
  return { changedRows: [...replaced.keys()].toSorted((a, b) => a - b) };
}

/** The read rule (see the module header): the sheet's rows with each hidden
 * row inside its autofilter's data rows reclassified. The same array when
 * nothing changes. */
export function classifyHiddenRows(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, ctx: FormatContext): Row[] {
  const filter = sheet.autoFilter;
  if (!filter) {
    return sheet.rows;
  }
  const evaluator = new FilterEvaluator(model, sheet, filter, ctx);
  const replaced = new Map<number, Row | undefined>();
  for (const row of sheet.rows) {
    if (!row.hidden || row.index <= filter.range.top || row.index > filter.range.bottom) {
      continue;
    }
    if (evaluator.hasPreserved || !evaluator.passes(row.index)) {
      const { hidden: _user, ...rest } = row;
      replaced.set(row.index, { ...rest, filterHidden: true });
    }
  }
  return withRows(sheet.rows, replaced);
}

/** The parts of `range` on rows not filter-hidden, as row spans (top to
 * bottom) — what copy, clear contents and fill act on (SPEC-128). */
export function visibleRowSpans(sheet: Sheet, range: GridRange): GridRange[] {
  const spans: GridRange[] = [];
  let top = range.top;
  for (const row of sheet.rows) {
    if (row.index < range.top || !row.filterHidden) {
      continue;
    }
    if (row.index > range.bottom) {
      break;
    }
    if (row.index > top) {
      spans.push({ ...range, top, bottom: row.index - 1 });
    }
    top = row.index + 1;
  }
  if (top <= range.bottom) {
    spans.push({ ...range, top, bottom: range.bottom });
  }
  return spans;
}

/** "Choosing the range" (see the module header): a single cell → its
 * currentRegion; a one-row selection → its columns down to the region's
 * bottom; a multi-row selection → itself. null: the region is empty. */
export function autoFilterRangeFor(selection: SheetSelection, nav: NavContext): GridRange | null {
  const active = selection.ranges[selection.activeRangeIndex] ?? selection.ranges[0];
  const at: CellRef = selection.active;
  const region = currentRegion(at, nav);
  const holdsContent = (range: GridRange): boolean =>
    nav.sheet.rows.some((row) => row.index >= range.top && row.index <= range.bottom && row.cells.some((cell) => cell.col >= range.left && cell.col <= range.right && cell.value !== null));
  if (!active || (active.top === active.bottom && active.left === active.right)) {
    return holdsContent(region) ? region : null;
  }
  if (active.top === active.bottom) {
    const range = { ...active, bottom: Math.max(active.bottom, region.bottom) };
    return holdsContent(range) ? range : null;
  }
  return { ...active };
}

// --- commands -------------------------------------------------------------------

const PROTECTED = (sheet: Sheet): boolean => sheet.protection !== undefined && !sheet.protection.allowAutoFilter;

function gridSheet(model: SpreadsheetModel, sheetId: string): Sheet | undefined {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  return sheet && sheet.kind === "grid" ? sheet : undefined;
}

/** Data ▸ Filter: off when the sheet has an autofilter (removing it and every
 * row's `filterHidden`), else on over `range` (autoFilterRangeFor's; null
 * refuses with EMPTY_FILTER_REGION_MESSAGE). */
export function toggleAutoFilter(model: SpreadsheetModel, sheetId: string, range: GridRange | null, ctx: FormatContext): FilterCommandResult {
  void ctx;
  const sheet = gridSheet(model, sheetId);
  if (!sheet) {
    return { ok: false, message: EMPTY_FILTER_REGION_MESSAGE };
  }
  if (PROTECTED(sheet)) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  if (sheet.autoFilter) {
    const { rows, changed } = clearedRows(sheet);
    replaceSheet(model, sheetId, (each) => {
      const { autoFilter: _gone, ...rest } = each;
      return { ...rest, rows, dirty: true };
    });
    model.structureDirty = true;
    return { ok: true, changedRows: changed };
  }
  if (!range) {
    return { ok: false, message: EMPTY_FILTER_REGION_MESSAGE };
  }
  replaceSheet(model, sheetId, (each) => ({ ...each, autoFilter: { range: { ...range }, columns: [] }, dirty: true }));
  model.structureDirty = true;
  return { ok: true, changedRows: [] };
}

/** A column's criteria set (undefined: cleared), then the filter applied. */
export function setFilterCriteria(
  model: SpreadsheetModel,
  sheetId: string,
  colId: number,
  criteria: FilterCriteria | undefined,
  ctx: FormatContext,
): FilterCommandResult {
  const sheet = gridSheet(model, sheetId);
  const filter = sheet?.autoFilter;
  if (!sheet || !filter) {
    return { ok: false, message: EMPTY_FILTER_REGION_MESSAGE };
  }
  if (PROTECTED(sheet)) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  const columns: FilterColumn[] = [];
  let placed = false;
  for (const column of filter.columns) {
    if (column.colId === colId) {
      placed = true;
      const next: FilterColumn = { colId };
      if (criteria) {
        next.criteria = criteria;
      }
      if (column.hideButton) {
        next.hideButton = true;
      }
      if (next.criteria || next.hideButton) {
        columns.push(next);
      }
    } else {
      columns.push(column);
    }
  }
  if (!placed && criteria) {
    columns.push({ colId, criteria });
    columns.sort((a, b) => a.colId - b.colId);
  }
  // An unmapped ods filter spans columns: setting any column's criteria
  // (clearing included) replaces it, so its rows show before applying.
  const { sourceXml: _source, odsUnmappedFilter, ...rest } = filter;
  const unmapped = odsUnmappedFilter === undefined ? { rows: sheet.rows, changed: [] } : clearedRows(sheet);
  replaceSheet(model, sheetId, (each) => ({ ...each, rows: unmapped.rows, autoFilter: { ...rest, columns }, dirty: true }));
  const { changedRows } = applyAutoFilter(model, sheetId, ctx);
  return { ok: true, changedRows: [...new Set([...unmapped.changed, ...changedRows])].toSorted((a, b) => a - b) };
}

/** Data ▸ Clear filters: every criterion removed (the buttons stay), every
 * row's `filterHidden` cleared; `hidden` never touched. */
export function clearFilters(model: SpreadsheetModel, sheetId: string): FilterCommandResult {
  const sheet = gridSheet(model, sheetId);
  const filter = sheet?.autoFilter;
  if (!sheet || !filter) {
    return { ok: false, message: EMPTY_FILTER_REGION_MESSAGE };
  }
  if (PROTECTED(sheet)) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  const { sourceXml: _source, odsUnmappedFilter: _unmapped, ...rest } = filter;
  const columns = filter.columns.filter((column) => column.hideButton).map((column): FilterColumn => ({ colId: column.colId, hideButton: true }));
  const { rows, changed } = clearedRows(sheet);
  replaceSheet(model, sheetId, (each) => ({ ...each, rows, autoFilter: { ...rest, columns }, dirty: true }));
  return { ok: true, changedRows: changed };
}

/** Data ▸ Reapply filter: the criteria re-evaluated against current values. */
export function reapplyFilter(model: SpreadsheetModel, sheetId: string, ctx: FormatContext): FilterCommandResult {
  const sheet = gridSheet(model, sheetId);
  if (!sheet?.autoFilter) {
    return { ok: false, message: EMPTY_FILTER_REGION_MESSAGE };
  }
  if (PROTECTED(sheet)) {
    return { ok: false, message: PROTECTED_CELL_MESSAGE };
  }
  const { changedRows } = applyAutoFilter(model, sheetId, ctx);
  replaceSheet(model, sheetId, (each) => (each.dirty ? each : { ...each, dirty: true }));
  return { ok: true, changedRows };
}

// --- what the UI reads -------------------------------------------------------------

/** The popover's value list: one item per distinct display text of the
 * column's data rows that pass every OTHER column's criteria (Excel's
 * list), sorted by the sort comparator; `blanks` null when no such row is
 * blank; past FILTER_VALUE_LIST_MAX the list is cut and `truncated`. */
export function filterValueList(
  model: Pick<SpreadsheetModel, "styles">,
  sheet: Sheet,
  colId: number,
  ctx: FormatContext,
): { items: { text: string; checked: boolean }[]; blanks: boolean | null; truncated: boolean } {
  const filter = sheet.autoFilter;
  if (!filter) {
    return { items: [], blanks: null, truncated: false };
  }
  const own = filter.columns.find((column) => column.colId === colId)?.criteria;
  const others: AutoFilter = { ...filter, columns: filter.columns.filter((column) => column.colId !== colId) };
  const otherEvaluator = new FilterEvaluator(model, sheet, others, ctx);
  const ownEvaluator = new FilterEvaluator(model, sheet, filter, ctx);
  const col = filter.range.left + colId;
  const seen = new Map<string, { text: string; value: CellValue; checked: boolean }>();
  let blanks: boolean | null = null;
  const last = lastDataRow(sheet, filter);
  for (let r = filter.range.top + 1; r <= last; r += 1) {
    if (!otherEvaluator.passes(r)) {
      continue;
    }
    const view = cellView(model, sheet, r, col, ctx);
    const passes = !own || own.kind === "preserved" || ownEvaluator.criteriaPass(colId, own, view);
    if (view.text === "") {
      blanks = (blanks ?? false) || passes;
      continue;
    }
    const entry = seen.get(view.text);
    if (entry) {
      entry.checked ||= passes;
    } else {
      seen.set(view.text, { text: displayText(model, sheet, r, col, ctx), value: view.value, checked: passes });
    }
  }
  const collator = sortCollator(ctx.locale.tag, false);
  const sorted = [...seen.values()].toSorted((a, b) => compareSortValues(a.value, b.value, collator) || collator.compare(a.text, b.text));
  const truncated = sorted.length > FILTER_VALUE_LIST_MAX;
  return {
    items: (truncated ? sorted.slice(0, FILTER_VALUE_LIST_MAX) : sorted).map((item) => ({ text: item.text, checked: item.checked })),
    blanks,
    truncated,
  };
}

/** A column's type for labels: the majority of its non-blank data cells —
 * date cells, other numbers, or everything else (ties → text); a formula
 * returning "" is blank here (display text ""), as in the criteria. `col` is
 * absolute; rows top..bottom inclusive. */
export function columnValueKind(
  model: Pick<SpreadsheetModel, "styles">,
  sheet: Sheet,
  col: number,
  rows: { top: number; bottom: number },
  ctx: FormatContext,
): "text" | "number" | "date" {
  let dates = 0;
  let numbers = 0;
  let texts = 0;
  for (const row of sheet.rows) {
    if (row.index < rows.top || row.index > rows.bottom) {
      continue;
    }
    const cell: Cell | undefined = row.cells.find((each) => each.col === col);
    if (!cell || cell.value === null || cell.value === "") {
      continue;
    }
    if (typeof cell.value === "number") {
      if (isDateCode(formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row.index, col)), ctx.locale))) {
        dates += 1;
      } else {
        numbers += 1;
      }
    } else {
      texts += 1;
    }
  }
  if (dates > numbers && dates > texts) {
    return "date";
  }
  return numbers > dates && numbers > texts ? "number" : "text";
}

/** "N of M records found" while the sheet's autofilter has a criterion: M
 * data rows up to the used range, N of them not filter-hidden; else null. */
export function filterRecordCounts(sheet: Sheet): { found: number; total: number } | null {
  const filter = sheet.autoFilter;
  if (!filter || !hasCriteria(filter)) {
    return null;
  }
  const last = lastDataRow(sheet, filter);
  const total = Math.max(0, last - filter.range.top);
  let hidden = 0;
  for (const row of sheet.rows) {
    if (row.filterHidden && row.index > filter.range.top && row.index <= last) {
      hidden += 1;
    }
  }
  return { found: total - hidden, total };
}

/** A header button's `data-state`: sorted (the filter's sort record's first
 * key on this column) wins over filtered. */
export function filterButtonState(filter: AutoFilter, colId: number): "none" | "filtered" | "sorted-asc" | "sorted-desc" {
  const key = filter.sort?.keys[0];
  if (key && key.col === filter.range.left + colId) {
    return key.descending ? "sorted-desc" : "sorted-asc";
  }
  if (filter.odsUnmappedFilter !== undefined || filter.columns.some((column) => column.colId === colId && column.criteria !== undefined)) {
    return "filtered";
  }
  return "none";
}
