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
import type { GridRange } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import type { CellKey } from "./formula/dependencies";
import type { CellValue, Sheet, SpreadsheetModel } from "./model";

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
  void col;
  throw new Error("not implemented");
}

export function sortCollator(localeTag: string, caseSensitive: boolean): Intl.Collator {
  void localeTag;
  void caseSensitive;
  throw new Error("not implemented");
}

/** Ascending order of two values (see the module header). */
export function compareSortValues(a: CellValue, b: CellValue, collator: Intl.Collator): number {
  void a;
  void b;
  void collator;
  throw new Error("not implemented");
}

export function guessHasHeader(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, range: GridRange): boolean {
  void model;
  void sheet;
  void range;
  throw new Error("not implemented");
}

export function sortRange(model: SpreadsheetModel, sheetId: string, range: GridRange, options: SortOptions): SortOutcome {
  void model;
  void sheetId;
  void range;
  void options;
  throw new Error("not implemented");
}

