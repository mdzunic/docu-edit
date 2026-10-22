// Find and replace, and Go To, over the workbook model (SPEC-134) — THIS FILE
// is the normative source in code for sheet search semantics, the sheet
// replace rules, the Go To reference grammar and Go To Special's selection
// rules. Framework-free: pure functions and one resumable walk over plain
// model arrays, mutating nothing, so searching never re-renders the grid.
// useSheetFind / useSheetGoTo hold the view state and apply the edits.
//
// - Search text ("Search text"): Values — a cell's SPEC-118 displayed text at
//   unlimited width (cellDisplay through its effective format; booleans
//   TRUE/FALSE, errors their code, rich text its plain `value`); Formulas —
//   "=" + `formula` for a formula cell (an array anchor included) and
//   SPEC-118's editText through the effective style for a constant. Never
//   searched: a cell storing no value (no value and no formula), an array
//   member in Formulas mode, a merge-covered cell — and, being outside the
//   cell grid, notes, comments, drawings, charts, names, foreign and chart
//   sheets, hidden and very hidden sheets. Hidden rows and columns are
//   searched.
// - Matching: compileCellQuery compiles through core's compileQuery (never
//   whole words), anchoring the source as `^(?:…)$` for "Match entire cell";
//   a cell matches when the query finds at least one NON-EMPTY match in its
//   text, and counts once however many it holds.
// - Scope and order: Within = Sheet searches the selection's ranges (in
//   `ranges` order, a cell an earlier range holds skipped) when the
//   selection covers more than one cell, else the whole active sheet; Within
//   = Workbook every visible grid sheet in tab order, whole. Row-major within
//   each. A hit's place in that order is its "search position";
//   `locate` answers at-or-after / after / before a cell, wrapping, or
//   "pending" when the walk has not got far enough to know.
// - The walk: `step(deadline)` visits stored cells only — a binary search to
//   each range's first stored row and, in each row, to its first cell ≥ the
//   range's left — and calls `deadline()` every FIND_CHECK_INTERVAL units of
//   work (a visited cell, or a stored row examined), returning as soon as it
//   is true, so no slot overruns its deadline by more than one interval.
// - The replace plan (`planReplace`, "Replace rules"), per hit in order: an
//   array anchor or member → `array`; a cell SPEC-122's protection rule
//   forbids (firstLockedCell) → `protected`; Values mode, a formula cell →
//   `formulaResult`; then every non-empty match of the search text is
//   replaced by the replacement verbatim (a replacer function: `$1`, `$&`
//   mean nothing); Formulas mode, a new text starting with "=" must parse
//   (parseFormula, within the 8,192-character limit) or → `breaksFormula`.
//   An edit's `asValue` (Values mode) means parseInput's text is stored as
//   typed but never as a formula. Pure: the caller enters the edits.
// - Go To references (`resolveGoToTarget`): the text splits on top-level
//   commas (tokenize: outside quotes and parentheses); each part is a cell,
//   range, whole-column or whole-row reference, optionally sheet-qualified,
//   case-insensitive and `$`-tolerant; a visible defined name
//   (resolveDefinedName from the active sheet, hidden and reserved names
//   never) whose value is a reference; or a table (its data body). Every
//   part lands on one grid sheet. 3-D, external and R1C1 references,
//   intersections, anything outside A1:XFD1048576, an unknown sheet or name
//   are GO_TO_INVALID_MESSAGE; a name that is no reference
//   GO_TO_NOT_A_RANGE_MESSAGE; a hidden sheet GO_TO_HIDDEN_SHEET_MESSAGE.
// - Go To Special (`selectSpecial`): within the selection's ranges (each
//   clipped to the used range) when it covers more than one cell, else the
//   used range; blanks as per-row complements of the stored columns (cost:
//   rows plus stored cells, never the area); constants, formulas (array
//   members included; "" is text, an uncomputed null a number) and errors by
//   type; the last cell; visible cells as the cross product of visible row
//   and column runs. Matching cells coalesce (`coalesceCells`) into maximal
//   horizontal runs, a run extending the rectangle directly above it with
//   the same left and right; more than GO_TO_SPECIAL_MAX_RANGES rectangles is
//   refused.
import { compileQuery, isQueryError, type QueryError } from "../core/findQuery";
import { findRow, getCell, usedRange, type CellRef, type GridRange } from "./address";
import { selectionSpansMultipleCells } from "./aggregates";
import { cellDisplay } from "./cellRender";
import { firstLockedCell } from "./cells";
import { findArrayAnchor } from "./formula/anchors";
import { storedValueContext } from "./formula/evaluate";
import { FormulaSyntaxError, parseFormula } from "./formula/parse";
import type { ReferenceNode } from "./formula/references";
import { evaluateDefinedName } from "./formula/referenceHooks";
import { joinTokens, tokenize, type Token } from "./formula/tokens";
import { isRange } from "./formula/values";
import { buildMergeIndex, mergeAt, type MergeIndex } from "./merges";
import { isCellError, MAX_COLS, MAX_ROWS, rowIsHidden, type Cell, type CellStyle, type DefinedName, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { isVisibleName, resolveDefinedName } from "./names";
import { editText, plainText, type FormatContext } from "./numberFormat/render";
import type { SheetSelection } from "./selection";
import { cellStyle, columnSpanAt } from "./styles";
import { dataBody, findTable } from "./tables";

export type { QueryError } from "../core/findQuery";

export type FindLookIn = "values" | "formulas";
export type FindWithin = "sheet" | "workbook";

export interface SheetFindOptions {
  matchCase: boolean;
  entireCell: boolean;
  regex: boolean;
  lookIn: FindLookIn;
  within: FindWithin;
}

export interface FindHit {
  sheetId: string;
  row: number;
  col: number;
}

export interface CompiledCellQuery {
  /** Non-global: the cheap "does it match at all" check. */
  test: RegExp;
  /** Global: every match, for the non-empty check and replace. */
  global: RegExp;
}

/** How often the walk asks its deadline, in units of work. */
export const FIND_CHECK_INTERVAL = 256;
/** A walk slot's time budget (useSheetFind's scheduleEagerWork slots). */
export const FIND_SLICE_BUDGET_MS = 8;
/** The pause after typing before the walk (re)starts. */
export const FIND_DEBOUNCE_MS = 150;
/** The most rows "Find all" lists. */
export const FIND_ALL_LIMIT = 5_000;
/** The most rectangles Go To Special selects. */
export const GO_TO_SPECIAL_MAX_RANGES = 10_000;

// --- messages ("Messages") ----------------------------------------------------------

export const FIND_SCOPE_LABEL = "Cells only";
export const FIND_SCOPE_TITLE = "Search covers cell values and formulas. Notes, comments, shapes, charts and hidden sheets are not searched.";
export const FIND_PLACEHOLDERS: Readonly<Record<FindWithin, string>> = { sheet: "Find in sheet", workbook: "Find in workbook" };
export const FIND_HIDDEN_MATCH_HINT = "This match is in a hidden row or column";
export const FIND_SEARCHING_HINT = "Searching…";
/** SPEC-017's wrap hint, shown when a step wraps past either end. */
export const FIND_WRAPPED_HINT = "Wrapped to top";
export const GO_TO_INVALID_MESSAGE = "That isn't a cell reference or a defined name";
export const GO_TO_HIDDEN_SHEET_MESSAGE = "That sheet is hidden";
/** SPEC-133's Name Box string, shared with Go To. */
export const GO_TO_NOT_A_RANGE_MESSAGE = "That name doesn't refer to a range";
export const GO_TO_SPECIAL_NONE_MESSAGE = "No cells were found.";
export const GO_TO_SPECIAL_TOO_MANY_MESSAGE = "Too many separate areas to select (more than 10,000)";

const plural = (count: number, one: string, many: string): string => `${count} ${count === 1 ? one : many}`;

/** A single Replace refused because the new text does not parse. */
export function breaksFormulaMessage(address: string): string {
  return `That replacement would break the formula in ${address}`;
}

/** "Showing the first 5,000 of N matches". */
export function findAllCapMessage(total: number): string {
  return `Showing the first ${FIND_ALL_LIMIT.toLocaleString("en-US")} of ${total.toLocaleString("en-US")} matches`;
}

/** "Replaced N occurrences in M cells", plus " · skipped " and the parts. */
export function replaceReport(plan: Pick<ReplacePlan, "occurrences" | "skipped">, cells: number): string {
  const head = `Replaced ${plural(plan.occurrences, "occurrence", "occurrences")} in ${plural(cells, "cell", "cells")}`;
  const { breaksFormula, protected: locked, array, formulaResult } = plan.skipped;
  const parts = [
    breaksFormula > 0 ? `${breaksFormula} that would break a formula` : "",
    locked > 0 ? `${locked} protected` : "",
    array > 0 ? `${array} in array formulas` : "",
    formulaResult > 0 ? plural(formulaResult, "formula result", "formula results") : "",
  ].filter((part) => part !== "");
  return parts.length === 0 ? head : `${head} · skipped ${parts.join(", ")}`;
}

// --- the query ------------------------------------------------------------------------

/** The query under the options (see the header); null for an empty query. */
export function compileCellQuery(query: string, options: Pick<SheetFindOptions, "matchCase" | "entireCell" | "regex">): CompiledCellQuery | QueryError | null {
  const compiled = compileQuery(query, { matchCase: options.matchCase, wholeWords: false, regex: options.regex });
  if (compiled === null || isQueryError(compiled)) {
    return compiled;
  }
  const source = options.entireCell ? `^(?:${compiled.regExp.source})$` : compiled.regExp.source;
  const flags = compiled.regExp.flags;
  return { test: new RegExp(source, flags.replace("g", "")), global: new RegExp(source, flags) };
}

/** Whether `text` holds at least one non-empty match. */
export function cellMatches(text: string, query: CompiledCellQuery): boolean {
  if (text === "" || !query.test.test(text)) {
    return false;
  }
  const global = query.global;
  global.lastIndex = 0;
  for (let match = global.exec(text); match !== null; match = global.exec(text)) {
    if (match[0] !== "") {
      global.lastIndex = 0;
      return true;
    }
    global.lastIndex += 1;
  }
  global.lastIndex = 0;
  return false;
}

/** `text` with every non-empty match replaced by `replacement` verbatim,
 * and how many were replaced. */
export function replaceInText(text: string, query: CompiledCellQuery, replacement: string): { text: string; count: number } {
  let count = 0;
  query.global.lastIndex = 0;
  const replaced = text.replace(query.global, (match: string) => {
    if (match === "") {
      return match;
    }
    count += 1;
    return replacement;
  });
  return { text: replaced, count };
}

// --- search text ----------------------------------------------------------------------

/** The index of the first item whose key is ≥ `at` (sorted ascending). */
function lowerBound<T>(items: readonly T[], at: number, key: (item: T) => number): number {
  let lo = 0;
  let hi = items.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (key(items[mid]!) < at) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

const rowIndex = (row: Row): number => row.index;
const cellCol = (cell: Cell): number => cell.col;

/** Per sheet (copy-on-write: a changed sheet is a new object), what every
 * cell's text needs once: its merge index and its resolved styles, the
 * latter kept while `model.styles` is the same array. */
interface SheetLookup {
  merges: MergeIndex;
  styleTable: readonly CellStyle[] | null;
  styles: Map<number, CellStyle>;
}

const lookups = new WeakMap<Sheet, SheetLookup>();

function lookupOf(sheet: Sheet): SheetLookup {
  let found = lookups.get(sheet);
  if (!found) {
    found = { merges: buildMergeIndex(sheet.merges), styleTable: null, styles: new Map() };
    lookups.set(sheet, found);
  }
  return found;
}

function styleOf(model: SpreadsheetModel, lookup: SheetLookup, styleId: number): CellStyle {
  if (lookup.styleTable !== model.styles) {
    lookup.styleTable = model.styles;
    lookup.styles.clear();
  }
  let style = lookup.styles.get(styleId);
  if (!style) {
    style = cellStyle(model, styleId);
    lookup.styles.set(styleId, style);
  }
  return style;
}

/** Whether (row, cell) is an array member: inside an array formula's range,
 * other than its anchor. */
function isArrayMember(sheet: Sheet, row: number, cell: Cell): boolean {
  if (cell.arrayRange !== undefined) {
    return false;
  }
  const anchor = findArrayAnchor(sheet, row, cell.col);
  const range = anchor?.arrayRange;
  return range !== undefined && !(range.top === row && range.left === cell.col);
}

function textOf(model: SpreadsheetModel, sheet: Sheet, row: Row, cell: Cell, lookIn: FindLookIn, ctx: FormatContext): string | null {
  if (cell.value === null && cell.formula === undefined) {
    return null;
  }
  const lookup = lookupOf(sheet);
  if (lookup.merges.ranges.length > 0) {
    const merge = mergeAt(lookup.merges, row.index, cell.col);
    if (merge && (merge.top !== row.index || merge.left !== cell.col)) {
      return null;
    }
  }
  const style = styleOf(model, lookup, cell.styleId ?? row.styleId ?? columnSpanAt(sheet.columns, cell.col)?.styleId ?? 0);
  if (lookIn === "formulas") {
    if (cell.formula !== undefined) {
      return `=${cell.formula}`;
    }
    return isArrayMember(sheet, row.index, cell) ? null : editText(cell, { ...ctx, style });
  }
  return plainText(cellDisplay(cell.value, style, sheet, { ...ctx, widthPx: undefined }));
}

/** A cell's search text under Look in (see the header); null when it is not
 * searched. `row` is the cell's row index — the walk always passes it; when
 * absent it is found by scanning the sheet's rows for the cell. */
export function searchText(model: SpreadsheetModel, sheet: Sheet, cell: Cell, lookIn: FindLookIn, ctx: FormatContext, row?: number): string | null {
  const host = row === undefined ? sheet.rows.find((each) => each.cells.includes(cell)) : findRow(sheet, row);
  return host ? textOf(model, sheet, host, cell, lookIn, ctx) : null;
}

// --- scope, order and the walk --------------------------------------------------------

const FULL_SHEET: GridRange = { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 };
/** Positions per area in the search order: one past the last cell. */
const AREA_SPAN = MAX_ROWS * MAX_COLS;

/** One stretch of the search order: a range of one sheet. */
export interface FindArea {
  sheet: Sheet;
  range: GridRange;
}

/** The scope (see the header), in search order. */
export function findAreas(model: SpreadsheetModel, within: FindWithin, activeSheetId: string, selection: SheetSelection): FindArea[] {
  if (within === "workbook") {
    return model.sheets.filter((sheet) => sheet.kind === "grid" && sheet.visibility === "visible").map((sheet) => ({ sheet, range: FULL_SHEET }));
  }
  const sheet = model.sheets.find((each) => each.id === activeSheetId);
  if (!sheet || sheet.kind !== "grid") {
    return [];
  }
  if (selectionSpansMultipleCells(selection.ranges, lookupOf(sheet).merges)) {
    return selection.ranges.map((range) => ({ sheet, range }));
  }
  return [{ sheet, range: FULL_SHEET }];
}

/** Where a request stands against the walk: a hit (and whether it wrapped
 * past an end), none at all, or not known until the walk goes further. */
export type FindLocation = { kind: "hit"; index: number; wrapped: boolean } | { kind: "none" } | { kind: "pending" };

export interface SheetSearch {
  /** In search order, appended as the walk proceeds. */
  readonly hits: FindHit[];
  readonly done: boolean;
  /** Returns when done or deadline() is true. */
  step(deadline: () => boolean): void;
  /** Units of work done so far (visited cells and examined rows). */
  readonly work: number;
  /** The first hit at or after (`atOrAfter`), after, or the last before a cell. */
  locate(at: FindHit, mode: "atOrAfter" | "after" | "before"): FindLocation;
}

const areasOverlap = (a: FindArea, b: FindArea): boolean =>
  a.sheet === b.sheet && a.range.top <= b.range.bottom && b.range.top <= a.range.bottom && a.range.left <= b.range.right && b.range.left <= a.range.right;
const byKey = (key: number): number => key;

const contains = (range: GridRange, row: number, col: number): boolean => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;

export function createSheetSearch(
  model: SpreadsheetModel,
  query: CompiledCellQuery,
  options: SheetFindOptions,
  activeSheetId: string,
  selection: SheetSelection,
  ctx: FormatContext,
): SheetSearch {
  const areas = findAreas(model, options.within, activeSheetId, selection);
  // Overlapping selection ranges: a cell counts in its first range only.
  const overlapping = areas.some((a, i) => areas.slice(0, i).some((b) => areasOverlap(a, b)));
  const hits: FindHit[] = [];
  const order: number[] = [];
  let area = 0;
  let rowAt = -1;
  let cellAt = -1;
  let work = 0;
  /** Every position ordered before this has been decided. */
  let decidedBefore = 0;
  let done = areas.length === 0;

  const key = (index: number, row: number, col: number): number => index * AREA_SPAN + row * MAX_COLS + col;

  function positionOf(at: FindHit): number {
    for (let index = 0; index < areas.length; index += 1) {
      const { sheet, range } = areas[index]!;
      if (sheet.id === at.sheetId && contains(range, at.row, at.col)) {
        return key(index, at.row, at.col);
      }
    }
    return -1;
  }

  function visit(index: number, sheet: Sheet, row: Row, cell: Cell): void {
    if (overlapping && areas.slice(0, index).some((earlier) => earlier.sheet === sheet && contains(earlier.range, row.index, cell.col))) {
      return;
    }
    const text = textOf(model, sheet, row, cell, options.lookIn, ctx);
    if (text !== null && cellMatches(text, query)) {
      hits.push({ sheetId: sheet.id, row: row.index, col: cell.col });
      order.push(key(index, row.index, cell.col));
    }
  }

  function step(deadline: () => boolean): void {
    let budget = FIND_CHECK_INTERVAL;
    const spend = (): boolean => {
      work += 1;
      budget -= 1;
      if (budget > 0) {
        return false;
      }
      budget = FIND_CHECK_INTERVAL;
      return deadline();
    };
    while (area < areas.length) {
      const { sheet, range } = areas[area]!;
      const rows = sheet.rows;
      if (rowAt < 0) {
        rowAt = lowerBound(rows, range.top, rowIndex);
        cellAt = -1;
      }
      while (rowAt < rows.length && rows[rowAt]!.index <= range.bottom) {
        const row = rows[rowAt]!;
        if (cellAt < 0) {
          cellAt = lowerBound(row.cells, range.left, cellCol);
          if (spend()) {
            decidedBefore = key(area, row.index, range.left);
            return;
          }
        }
        while (cellAt < row.cells.length && row.cells[cellAt]!.col <= range.right) {
          const cell = row.cells[cellAt]!;
          visit(area, sheet, row, cell);
          cellAt += 1;
          if (spend()) {
            decidedBefore = key(area, row.index, cell.col + 1);
            return;
          }
        }
        rowAt += 1;
        cellAt = -1;
        decidedBefore = key(area, row.index + 1, 0);
      }
      area += 1;
      rowAt = -1;
      decidedBefore = key(area, 0, 0);
    }
    done = true;
  }

  function locate(at: FindHit, mode: "atOrAfter" | "after" | "before"): FindLocation {
    const position = positionOf(at);
    if (mode !== "before") {
      const index = lowerBound(order, mode === "after" ? position + 1 : position, byKey);
      if (index < hits.length) {
        return { kind: "hit", index, wrapped: false };
      }
      if (!done) {
        return { kind: "pending" };
      }
      return hits.length === 0 ? { kind: "none" } : { kind: "hit", index: 0, wrapped: mode === "after" && position >= 0 };
    }
    if (!done && (position < 0 || decidedBefore < position)) {
      return { kind: "pending" };
    }
    const index = position < 0 ? -1 : lowerBound(order, position, byKey) - 1;
    if (index >= 0) {
      return { kind: "hit", index, wrapped: false };
    }
    if (!done) {
      return { kind: "pending" };
    }
    return hits.length === 0 ? { kind: "none" } : { kind: "hit", index: hits.length - 1, wrapped: true };
  }

  return {
    hits,
    get done() {
      return done;
    },
    get work() {
      return work;
    },
    step,
    locate,
  };
}

/** Whether a hit sits in a hidden row (the user's or a filter's) or column. */
export function hitIsHidden(model: SpreadsheetModel, hit: FindHit): boolean {
  const sheet = model.sheets.find((each) => each.id === hit.sheetId);
  return sheet !== undefined && (rowIsHidden(findRow(sheet, hit.row)) || columnSpanAt(sheet.columns, hit.col)?.hidden === true);
}

// --- the replace plan -----------------------------------------------------------------

export interface ReplacePlan {
  edits: { sheetId: string; row: number; col: number; text: string; asValue: boolean }[];
  occurrences: number;
  skipped: { breaksFormula: number; protected: number; array: number; formulaResult: number };
}

/** SPEC-122's MAX_FORMULA_LENGTH (Excel's limit), "=" excluded. */
const MAX_FORMULA_LENGTH = 8192;

function parses(body: string): boolean {
  if (body.length > MAX_FORMULA_LENGTH) {
    return false;
  }
  try {
    parseFormula(body);
    return true;
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return false;
    }
    throw error;
  }
}

/** The edits "Replace" / "Replace all" make (see the header). Pure. */
export function planReplace(
  model: SpreadsheetModel,
  hits: readonly FindHit[],
  query: CompiledCellQuery,
  replacement: string,
  options: Pick<SheetFindOptions, "lookIn">,
  ctx: FormatContext,
): ReplacePlan {
  const plan: ReplacePlan = { edits: [], occurrences: 0, skipped: { breaksFormula: 0, protected: 0, array: 0, formulaResult: 0 } };
  const indexOf = new Map(model.sheets.map((sheet, index) => [sheet.id, index]));
  for (const hit of hits) {
    const sheetIndex = indexOf.get(hit.sheetId);
    const sheet = sheetIndex === undefined ? undefined : model.sheets[sheetIndex];
    const row = sheet ? findRow(sheet, hit.row) : undefined;
    const cell = sheet ? getCell(sheet, hit.row, hit.col) : undefined;
    if (sheetIndex === undefined || !sheet || !row || !cell) {
      continue;
    }
    if (cell.arrayRange !== undefined || isArrayMember(sheet, hit.row, cell)) {
      plan.skipped.array += 1;
      continue;
    }
    if (firstLockedCell(model, sheetIndex, [{ top: hit.row, left: hit.col, bottom: hit.row, right: hit.col }]) !== null) {
      plan.skipped.protected += 1;
      continue;
    }
    if (options.lookIn === "values" && cell.formula !== undefined) {
      plan.skipped.formulaResult += 1;
      continue;
    }
    const text = textOf(model, sheet, row, cell, options.lookIn, ctx);
    const replaced = text === null ? null : replaceInText(text, query, replacement);
    if (!replaced || replaced.count === 0) {
      continue;
    }
    if (options.lookIn === "formulas" && replaced.text.startsWith("=") && !parses(replaced.text.slice(1))) {
      plan.skipped.breaksFormula += 1;
      continue;
    }
    plan.edits.push({ sheetId: hit.sheetId, row: hit.row, col: hit.col, text: replaced.text, asValue: options.lookIn === "values" });
    plan.occurrences += replaced.count;
  }
  return plan;
}

// --- Go To references -------------------------------------------------------------------

export type GoToTarget = { sheetId: string; ranges: GridRange[] };

const invalid = (): { error: string } => ({ error: GO_TO_INVALID_MESSAGE });

/** The text's parts, split on top-level commas (see the header). */
function goToParts(text: string): string[] {
  const parts: Token[][] = [[]];
  let depth = 0;
  for (const token of tokenize(text)) {
    if (token.kind === "open") {
      depth += 1;
    } else if (token.kind === "close") {
      depth -= 1;
    }
    if (token.kind === "union" && depth === 0) {
      parts.push([]);
    } else {
      parts.at(-1)!.push(token);
    }
  }
  return parts.map((tokens) => joinTokens(tokens).trim());
}

function referenceRange(ref: ReferenceNode): GridRange {
  const end = ref.end ?? ref.start;
  const wholeColumns = ref.start.row === undefined || end.row === undefined;
  const wholeRows = ref.start.col === undefined || end.col === undefined;
  return {
    top: wholeColumns ? 0 : Math.min(ref.start.row!, end.row!),
    left: wholeRows ? 0 : Math.min(ref.start.col!, end.col!),
    bottom: wholeColumns ? MAX_ROWS - 1 : Math.max(ref.start.row!, end.row!),
    right: wholeRows ? MAX_COLS - 1 : Math.max(ref.start.col!, end.col!),
  };
}

/** Where a visible name selects: its evaluation at `at` on the active sheet
 * (its scope sheet's for a local name) — the first area's sheet, its areas. */
function nameTarget(model: SpreadsheetModel, name: DefinedName, activeSheetId: string, at: CellRef): GoToTarget | { error: string } {
  const value = evaluateDefinedName(storedValueContext(model), name, { sheetId: name.scopeSheetId ?? activeSheetId, row: at.row, col: at.col });
  const sheetId = isRange(value) ? value.areas[0]?.sheetId : undefined;
  if (!isRange(value) || sheetId === undefined || model.sheets.find((sheet) => sheet.id === sheetId)?.kind !== "grid") {
    return { error: GO_TO_NOT_A_RANGE_MESSAGE };
  }
  return { sheetId, ranges: value.areas.filter((area) => area.sheetId === sheetId).map(({ top, left, bottom, right }) => ({ top, left, bottom, right })) };
}

function partTarget(model: SpreadsheetModel, part: string, activeSheetId: string, at: CellRef): GoToTarget | { error: string } {
  let node: ReturnType<typeof parseFormula>;
  try {
    node = parseFormula(part);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return invalid();
    }
    throw error;
  }
  if (node.kind === "ref") {
    const qualifier = node.ref.sheet;
    if (node.ref.sheetEnd !== undefined) {
      return invalid();
    }
    const sheet = qualifier === undefined ? model.sheets.find((each) => each.id === activeSheetId) : model.sheets.find((each) => each.name.toUpperCase() === qualifier.toUpperCase());
    return sheet && sheet.kind === "grid" ? { sheetId: sheet.id, ranges: [referenceRange(node.ref)] } : invalid();
  }
  if (node.kind === "name") {
    const name = resolveDefinedName(model, node.name, node.sheet, activeSheetId);
    if (name && isVisibleName(name)) {
      return nameTarget(model, name, activeSheetId, at);
    }
    const table = node.sheet === undefined ? findTable(model, node.name) : undefined;
    if (table) {
      return { sheetId: table.sheet.id, ranges: [dataBody(table.table)] };
    }
  }
  return invalid();
}

/** What a Go To (or Name Box) entry selects (see the header). `at` is the
 * active cell, where a relative name evaluates (default A1). */
export function resolveGoToTarget(model: SpreadsheetModel, text: string, activeSheetId: string, at: CellRef = { row: 0, col: 0 }): GoToTarget | { error: string } {
  const parts = goToParts(text);
  if (parts.some((part) => part === "")) {
    return invalid();
  }
  let target: GoToTarget | null = null;
  for (const part of parts) {
    const found = partTarget(model, part, activeSheetId, at);
    if ("error" in found) {
      return found;
    }
    if (target && target.sheetId !== found.sheetId) {
      return invalid();
    }
    target = target ? { sheetId: target.sheetId, ranges: [...target.ranges, ...found.ranges] } : found;
  }
  const sheet = model.sheets.find((each) => each.id === target?.sheetId);
  if (!target || !sheet) {
    return invalid();
  }
  return sheet.visibility === "visible" ? target : { error: GO_TO_HIDDEN_SHEET_MESSAGE };
}

// --- Go To Special ----------------------------------------------------------------------

export type SpecialKind = "blanks" | "constants" | "formulas" | "errors" | "lastCell" | "visible";

export interface SpecialTypes {
  numbers: boolean;
  text: boolean;
  logicals: boolean;
  errors: boolean;
}

/** Rectangles built row by row from horizontal runs (see the header). */
class Coalescer {
  readonly out: GridRange[] = [];
  private open = new Map<string, GridRange>();
  private next = new Map<string, GridRange>();
  private row = -1;

  constructor(private readonly limit: number) {}

  get overflow(): boolean {
    return this.out.length > this.limit;
  }

  /** A run on `row` (rows ascending, runs in a row ascending and disjoint). */
  add(row: number, left: number, right: number): void {
    if (row !== this.row) {
      this.open = row === this.row + 1 ? this.next : new Map();
      this.next = new Map();
      this.row = row;
    }
    const runKey = `${left}:${right}`;
    const above = this.open.get(runKey);
    if (above && above.bottom === row - 1) {
      above.bottom = row;
      this.next.set(runKey, above);
      return;
    }
    const rect = { top: row, left, bottom: row, right };
    this.out.push(rect);
    this.next.set(runKey, rect);
  }
}

/** Cells as maximal rectangles (see the header): sorted row-major and
 * deduplicated first. */
export function coalesceCells(cells: Iterable<CellRef>): GridRange[] {
  const sorted = [...cells].toSorted((a, b) => a.row - b.row || a.col - b.col);
  const coalescer = new Coalescer(Number.POSITIVE_INFINITY);
  let run: { row: number; left: number; right: number } | null = null;
  for (const { row, col } of sorted) {
    if (run && run.row === row && col <= run.right + 1) {
      run.right = Math.max(run.right, col);
      continue;
    }
    if (run) {
      coalescer.add(run.row, run.left, run.right);
    }
    run = { row, left: col, right: col };
  }
  if (run) {
    coalescer.add(run.row, run.left, run.right);
  }
  return coalescer.out;
}

function intersect(a: GridRange, b: GridRange): GridRange | null {
  const top = Math.max(a.top, b.top);
  const left = Math.max(a.left, b.left);
  const bottom = Math.min(a.bottom, b.bottom);
  const right = Math.min(a.right, b.right);
  return top <= bottom && left <= right ? { top, left, bottom, right } : null;
}

/** The value kind a type checkbox filters. */
function typeOf(value: Cell["value"]): keyof SpecialTypes {
  if (value === null || typeof value === "number") {
    return "numbers";
  }
  if (typeof value === "string") {
    return "text";
  }
  return typeof value === "boolean" ? "logicals" : "errors";
}

/** Each stored cell of `range`, row-major. */
function* storedCells(sheet: Sheet, range: GridRange): Generator<{ row: number; cell: Cell }> {
  const rows = sheet.rows;
  for (let r = lowerBound(rows, range.top, rowIndex); r < rows.length && rows[r]!.index <= range.bottom; r += 1) {
    const row = rows[r]!;
    for (let c = lowerBound(row.cells, range.left, cellCol); c < row.cells.length && row.cells[c]!.col <= range.right; c += 1) {
      yield { row: row.index, cell: row.cells[c]! };
    }
  }
}

const hasValue = (cell: Cell): boolean => cell.value !== null || cell.formula !== undefined;

/** The blanks of `range` as per-row runs, fed to `coalescer`. */
function blankRuns(sheet: Sheet, range: GridRange, coalescer: Coalescer): void {
  const rows = sheet.rows;
  let r = lowerBound(rows, range.top, rowIndex);
  for (let row = range.top; row <= range.bottom && !coalescer.overflow; row += 1) {
    const stored = rows[r]?.index === row ? rows[r++] : undefined;
    let next = range.left;
    if (stored) {
      for (let c = lowerBound(stored.cells, range.left, cellCol); c < stored.cells.length && stored.cells[c]!.col <= range.right; c += 1) {
        const cell = stored.cells[c]!;
        if (!hasValue(cell)) {
          continue;
        }
        if (cell.col > next) {
          coalescer.add(row, next, cell.col - 1);
        }
        next = cell.col + 1;
      }
    }
    if (next <= range.right) {
      coalescer.add(row, next, range.right);
    }
  }
}

/** The visible runs of [from, to] given what is hidden, ascending. */
function visibleRuns(from: number, to: number, hidden: Iterable<{ min: number; max: number }>): { min: number; max: number }[] {
  const runs: { min: number; max: number }[] = [];
  let next = from;
  for (const span of hidden) {
    if (span.max < from || span.min > to) {
      continue;
    }
    if (span.min > next) {
      runs.push({ min: next, max: span.min - 1 });
    }
    next = Math.max(next, span.max + 1);
  }
  if (next <= to) {
    runs.push({ min: next, max: to });
  }
  return runs;
}

function* hiddenRows(sheet: Sheet, range: GridRange): Generator<{ min: number; max: number }> {
  const rows = sheet.rows;
  for (let r = lowerBound(rows, range.top, rowIndex); r < rows.length && rows[r]!.index <= range.bottom; r += 1) {
    if (rowIsHidden(rows[r])) {
      yield { min: rows[r]!.index, max: rows[r]!.index };
    }
  }
}

/** Go To Special's selection (see the header), or why there is none. */
export function selectSpecial(model: SpreadsheetModel, sheet: Sheet, selection: SheetSelection, kind: SpecialKind, types: SpecialTypes): GridRange[] | { error: string } {
  const used = usedRange(sheet);
  if (kind === "lastCell") {
    const row = used?.bottom ?? 0;
    const col = used?.right ?? 0;
    return [{ top: row, left: col, bottom: row, right: col }];
  }
  if (!used) {
    return { error: GO_TO_SPECIAL_NONE_MESSAGE };
  }
  const scope = selectionSpansMultipleCells(selection.ranges, lookupOf(sheet).merges)
    ? selection.ranges.map((range) => intersect(range, used)).filter((range): range is GridRange => range !== null)
    : [used];
  let rects: GridRange[] = [];
  if (kind === "blanks" || kind === "visible") {
    for (const range of scope) {
      if (kind === "blanks") {
        const coalescer = new Coalescer(GO_TO_SPECIAL_MAX_RANGES);
        blankRuns(sheet, range, coalescer);
        rects.push(...coalescer.out);
      } else {
        const columns = visibleRuns(range.left, range.right, sheet.columns.filter((span) => span.hidden));
        for (const rowRun of visibleRuns(range.top, range.bottom, hiddenRows(sheet, range))) {
          rects.push(...columns.map((column) => ({ top: rowRun.min, left: column.min, bottom: rowRun.max, right: column.max })));
          if (rects.length > GO_TO_SPECIAL_MAX_RANGES) {
            break;
          }
        }
      }
      if (rects.length > GO_TO_SPECIAL_MAX_RANGES) {
        break;
      }
    }
  } else {
    const cells: CellRef[] = [];
    for (const range of scope) {
      for (const { row, cell } of storedCells(sheet, range)) {
        const member = cell.formula === undefined && isArrayMember(sheet, row, cell);
        const matches =
          kind === "errors"
            ? isCellError(cell.value)
            : kind === "formulas"
              ? (cell.formula !== undefined || member) && types[typeOf(cell.value)]
              : cell.value !== null && cell.formula === undefined && !member && types[typeOf(cell.value)];
        if (matches) {
          cells.push({ row, col: cell.col });
        }
      }
    }
    rects = coalesceCells(cells);
  }
  if (rects.length > GO_TO_SPECIAL_MAX_RANGES) {
    return { error: GO_TO_SPECIAL_TOO_MANY_MESSAGE };
  }
  return rects.length === 0 ? { error: GO_TO_SPECIAL_NONE_MESSAGE } : rects;
}
