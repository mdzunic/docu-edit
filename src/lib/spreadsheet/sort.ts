// Sorting a sheet's rows (SPEC-128 "Sort semantics") — THIS FILE is the
// normative source for the spreadsheet sort: the type order, collation,
// blanks, stability, what moves and what refuses, the header guess and the
// sort record. SPEC-076's lib/tableSort.ts sorts the word processor's table
// rows as text with numeric collation; nothing here imports or moves it.
//
// - Classes. A blank (`value === null`, formula or not) is outside the
//   order. Ascending: number < string < FALSE < TRUE < error; numbers
//   numerically, strings through sortCollator, every error equal (dates are
//   numbers). A descending key negates its class and in-class comparison;
//   blanks then go after every non-blank, so they are last both ways.
// - Collator. `Intl.Collator(tag, { usage: "sort", numeric: false,
//   ignorePunctuation: false, sensitivity: cs ? "variant" : "accent",
//   caseFirst: "lower" })`, built once per sort — "A10" < "A2", "apple" and
//   "Apple" tie unless case-sensitive, where "apple" comes first.
// - Stability. Each record is decorated with its original index; the
//   comparator falls through the keys and then to that index.
// - Records. Data rows are the range's rows less the header row when
//   `hasHeader`; the visible positions P are those neither `hidden` nor
//   `filterHidden`. Each position contributes the record of its cells in
//   [left, right] (an absent row or cell is a blank). Sorted records are
//   written back into P in order: every Cell field travels; a formula moving
//   from row r to r′ is shifted like a copy (shiftFormula by r′ − r — Excel's
//   semantics, so a relative reference to another row moves with it), a
//   `formulaUnparsed` one verbatim, a single-row `arrayRange` with its
//   anchor. Row properties stay with the row index; formulas outside the
//   range that point into it are not adjusted. New Row objects are built only
//   for rows whose cells changed.
// - Permutation hook. Afterwards every ROW_PERMUTATION_ADJUSTERS entry runs,
//   in registry order, with the sheet, the range and the permutation (each
//   moved data row's `from` → `to`) under SPEC-124's copy-on-write rule —
//   the seam per-cell structures that travel with a record register into
//   (SPEC-135's annotations). Range-based structures never register.
// - Refusals change nothing and carry SORT_REFUSAL_MESSAGES' text.
import { colToLetters, findRow, getCell, usedRange, type GridRange } from "./address";
import { isEmptyRow, PROTECTED_CELL_MESSAGE } from "./cells";
import { cellKey, type CellKey } from "./formula/dependencies";
import { shiftFormula } from "./formula/references";
import { isCellError, rowIsHidden, type Cell, type CellValue, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { withRows } from "./sheetEdit";
import { cellStyle, effectiveStyleId } from "./styles";

export const MAX_SORT_KEYS = 64;

/** One level of a sort: an absolute 0-based column and its direction. */
export interface SheetSortKey {
  col: number;
  descending: boolean;
}

export interface SortOptions {
  keys: SheetSortKey[];
  hasHeader: boolean;
  caseSensitive: boolean;
  localeTag: string;
}

export type SortRefusal = "merges" | "array" | "multiArea" | "tooFewRows" | "duplicateKey" | "protected";

export type SortOutcome =
  | { ok: true; permutation: { from: number; to: number }[]; changed: CellKey[] }
  | { ok: false; reason: SortRefusal; message: string };

/** The record of the last sort (SPEC-128 "The sort record") — into
 * `autoFilter.sort` when the range was an autofilter's data rows, else into
 * `Sheet.sortState`. A record only: never re-applied on open or on edit. */
export interface SortState {
  ref: GridRange; // the sorted data rows (header excluded)
  keys: SheetSortKey[]; // level order
  caseSensitive?: true;
  /** The source element verbatim while unchanged (keeps `sortBy="cellColor"`,
   *  `customList`, `columnSort`, which are no keys). */
  sourceXml?: string;
}

/** Per-cell structures that travel with a sorted record register here. */
export interface RowPermutationAdjuster {
  id: string;
  permute(model: SpreadsheetModel, sheetId: string, range: GridRange, permutation: readonly { from: number; to: number }[]): void; // absolute row indices
}

/** Empty in SPEC-128; SPEC-135 registers its annotations. */
export const ROW_PERMUTATION_ADJUSTERS: RowPermutationAdjuster[] = [];

export const SORT_REFUSAL_MESSAGES: Readonly<Record<Exclude<SortRefusal, "duplicateKey">, string>> = {
  merges: "Sorting isn't possible here — the range contains merged cells that span rows. Unmerge them first.",
  array: "Sorting would split an array formula — the range contains part of one.",
  multiArea: "Sorting works on one range at a time — select a single range.",
  tooFewRows: "There's nothing to sort — select a cell inside a range of at least two rows.",
  protected: PROTECTED_CELL_MESSAGE,
};

/** "Column B is used by more than one level — remove the duplicate." */
export function duplicateKeyMessage(col: number): string {
  return `Column ${colToLetters(col)} is used by more than one level — remove the duplicate.`;
}

export function sortCollator(localeTag: string, caseSensitive: boolean): Intl.Collator {
  return new Intl.Collator(localeTag, {
    usage: "sort",
    numeric: false,
    ignorePunctuation: false,
    sensitivity: caseSensitive ? "variant" : "accent",
    caseFirst: "lower",
  });
}

/** A value's class in ascending order; a blank (null) after every class. */
function classOf(value: CellValue): number {
  if (value === null) {
    return 5;
  }
  if (typeof value === "number") {
    return 0;
  }
  if (typeof value === "string") {
    return 1;
  }
  if (typeof value === "boolean") {
    return value ? 3 : 2;
  }
  return isCellError(value) ? 4 : 5;
}

/** Ascending order of two values (see the module header): numbers < text <
 * FALSE < TRUE < errors, a blank after everything. */
export function compareSortValues(a: CellValue, b: CellValue, collator: Intl.Collator): number {
  const ca = classOf(a);
  const cb = classOf(b);
  if (ca !== cb) {
    return ca - cb;
  }
  if (typeof a === "number" && typeof b === "number") {
    return a < b ? -1 : a > b ? 1 : 0;
  }
  if (typeof a === "string" && typeof b === "string") {
    return collator.compare(a, b);
  }
  return 0;
}

/** One key's comparison: blanks last in both directions. */
function compareByKey(a: CellValue, b: CellValue, descending: boolean, collator: Intl.Collator): number {
  if (a === null || b === null) {
    return a === null ? (b === null ? 0 : 1) : -1;
  }
  const order = compareSortValues(a, b, collator);
  return descending ? -order : order;
}

/** The stored cells of `row` in [left, right]. */
function cellsIn(row: Row | undefined, left: number, right: number): Cell[] {
  return row ? row.cells.filter((cell) => cell.col >= left && cell.col <= right) : [];
}

/** Whether the position resolves bold (SPEC-117's cell → row → column style). */
function isBold(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, row: number, col: number): boolean {
  return cellStyle(model, effectiveStyleId(sheet, row, col)).font?.bold === true;
}

/** "The header guess" (SPEC-128): pinned, and only ever the seed of the
 * "My data has headers" checkbox. */
export function guessHasHeader(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, range: GridRange): boolean {
  if (range.bottom - range.top + 1 < 2) {
    return false;
  }
  if (sheet.autoFilter && sheet.autoFilter.range.top === range.top) {
    return true;
  }
  const first = cellsIn(findRow(sheet, range.top), range.left, range.right).filter((cell) => cell.value !== null);
  if (first.length === 0 || first.some((cell) => typeof cell.value !== "string")) {
    return false;
  }
  const typedBelow = first.some((cell) => {
    const below = getCell(sheet, range.top + 1, cell.col)?.value ?? null;
    return below !== null && typeof below !== "string";
  });
  if (typedBelow) {
    return true;
  }
  const second = cellsIn(findRow(sheet, range.top + 1), range.left, range.right).filter((cell) => cell.value !== null);
  return first.every((cell) => isBold(model, sheet, range.top, cell.col)) && !second.some((cell) => isBold(model, sheet, range.top + 1, cell.col));
}

const refuse = (reason: Exclude<SortRefusal, "duplicateKey">): SortOutcome => ({ ok: false, reason, message: SORT_REFUSAL_MESSAGES[reason] });

const intersects = (a: GridRange, b: GridRange): boolean => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;

/** Why the records of `data` cannot move, or null: a merge spanning rows or
 * crossing the left/right edge, part of a multi-row (or edge-crossing) array. */
function blockedBy(sheet: Sheet, data: GridRange): "merges" | "array" | null {
  for (const merge of sheet.merges) {
    if (intersects(merge, data) && (merge.top !== merge.bottom || merge.left < data.left || merge.right > data.right)) {
      return "merges";
    }
  }
  for (const row of sheet.rows) {
    if (row.index > data.bottom) {
      break;
    }
    for (const cell of row.cells) {
      const array = cell.arrayRange;
      if (array && intersects(array, data) && (array.top !== array.bottom || array.left < data.left || array.right > data.right)) {
        return "array";
      }
    }
  }
  return null;
}

/** A record's cells moved by `delta` rows, shifted like a copy. */
function movedCells(cells: readonly Cell[], delta: number): Cell[] {
  return cells.map((cell) => {
    if (delta === 0 || (cell.formula === undefined && cell.arrayRange === undefined)) {
      return cell;
    }
    const next: Cell = { ...cell };
    if (next.formula !== undefined && !next.formulaUnparsed) {
      next.formula = shiftFormula(next.formula, delta, 0);
    }
    if (next.arrayRange) {
      next.arrayRange = { ...next.arrayRange, top: next.arrayRange.top + delta, bottom: next.arrayRange.bottom + delta };
    }
    return next;
  });
}

/** `row`'s cells with those in [left, right] replaced by `cells` (sorted). */
function spliceCells(row: Row | undefined, left: number, right: number, cells: readonly Cell[]): Cell[] {
  const kept = row ? row.cells.filter((cell) => cell.col < left || cell.col > right) : [];
  return [...kept, ...cells].toSorted((a, b) => a.col - b.col);
}

/** Sorts the rows of `range` on `sheetId` (see the module header) and
 * records the sort. A list of ranges sorts only when it holds exactly one. */
export function sortRange(model: SpreadsheetModel, sheetId: string, range: GridRange | readonly GridRange[], options: SortOptions): SortOutcome {
  const ranges: readonly GridRange[] = "top" in range ? [range] : range;
  if (ranges.length !== 1) {
    return refuse("multiArea");
  }
  const target = ranges[0]!;
  const sheetIndex = model.sheets.findIndex((each) => each.id === sheetId);
  const sheet = model.sheets[sheetIndex];
  if (!sheet || sheet.kind !== "grid") {
    return refuse("tooFewRows");
  }
  if (sheet.protection && !sheet.protection.allowSort) {
    return refuse("protected");
  }
  const dataTop = target.top + (options.hasHeader ? 1 : 0);
  if (target.bottom - dataTop + 1 < 2) {
    return refuse("tooFewRows");
  }
  const seen = new Set<number>();
  for (const key of options.keys) {
    if (seen.has(key.col)) {
      return { ok: false, reason: "duplicateKey", message: duplicateKeyMessage(key.col) };
    }
    seen.add(key.col);
  }
  const data: GridRange = { top: dataTop, left: target.left, bottom: target.bottom, right: target.right };
  const blocked = blockedBy(sheet, data);
  if (blocked) {
    return refuse(blocked);
  }

  // Rows past the used range hold only blanks, which sort last: they never move.
  const used = usedRange(sheet);
  const lastRow = Math.min(data.bottom, used?.bottom ?? -1);
  const positions: number[] = [];
  const records: { index: number; from: number; cells: Cell[] }[] = [];
  for (let r = data.top; r <= lastRow; r += 1) {
    const row = findRow(sheet, r);
    if (rowIsHidden(row)) {
      continue;
    }
    positions.push(r);
    records.push({ index: records.length, from: r, cells: cellsIn(row, data.left, data.right) });
  }
  const collator = sortCollator(options.localeTag, options.caseSensitive);
  const keyed = records.map((record) => ({
    record,
    values: options.keys.map((key) => record.cells.find((cell) => cell.col === key.col)?.value ?? null),
  }));
  keyed.sort((a, b) => {
    for (let k = 0; k < options.keys.length; k += 1) {
      const order = compareByKey(a.values[k]!, b.values[k]!, options.keys[k]!.descending, collator);
      if (order !== 0) {
        return order;
      }
    }
    return a.record.index - b.record.index;
  });

  const permutation: { from: number; to: number }[] = [];
  const changed: CellKey[] = [];
  const replaced = new Map<number, Row | undefined>();
  const rowMoves = new Map<number, number>();
  keyed.forEach(({ record }, i) => {
    const to = positions[i]!;
    if (record.from === to) {
      return;
    }
    permutation.push({ from: record.from, to });
    rowMoves.set(record.from, to);
    const destination = findRow(sheet, to);
    const before = cellsIn(destination, data.left, data.right);
    const cells = movedCells(record.cells, to - record.from);
    const next: Row = { ...(destination ?? { index: to }), index: to, cells: spliceCells(destination, data.left, data.right, cells) };
    replaced.set(to, isEmptyRow(next) ? undefined : next);
    for (const col of new Set([...before.map((cell) => cell.col), ...cells.map((cell) => cell.col)])) {
      changed.push(cellKey(sheetId, to, col));
    }
  });

  let merges = sheet.merges;
  if (rowMoves.size > 0 && merges.some((merge) => rowMoves.has(merge.top) && intersects(merge, data))) {
    merges = merges.map((merge) => {
      const to = intersects(merge, data) ? rowMoves.get(merge.top) : undefined;
      return to === undefined ? merge : { ...merge, top: to, bottom: to };
    });
  }
  const state: SortState = { ref: data, keys: options.keys.map((key) => ({ ...key })) };
  if (options.caseSensitive) {
    state.caseSensitive = true;
  }
  const filter = sheet.autoFilter;
  const updated: Sheet = { ...sheet, rows: withRows(sheet.rows, replaced), merges, dirty: true };
  if (
    filter &&
    filter.range.top + 1 === data.top &&
    filter.range.bottom === data.bottom &&
    filter.range.left === data.left &&
    filter.range.right === data.right
  ) {
    const { sourceXml: _source, ...rest } = filter;
    updated.autoFilter = { ...rest, sort: state };
  } else {
    updated.sortState = state;
  }
  model.sheets = model.sheets.map((each, i) => (i === sheetIndex ? updated : each));
  for (const adjuster of ROW_PERMUTATION_ADJUSTERS) {
    adjuster.permute(model, sheetId, data, permutation);
  }
  return { ok: true, permutation, changed };
}
