// What the sort and filter surfaces read (SPEC-128 "UI"), framework-free:
// a header's name, the Sort dialog's column and order choices, and the
// filter popover's value entries and the criterion its OK writes.
//
// - A column is named by its header cell's display text (its value through
//   its number format, filter.ts's rule) — "Column B" when that is blank or
//   when the data has no header row.
// - Order choices read "A to Z" / "Z to A", or "Smallest to Largest" /
//   "Largest to Smallest" for a column whose non-blank data cells are mostly
//   numbers or dates (columnValueKind; ties → text).
// - The popover lists filterValueList's items, then "Date groups" while the
//   column holds date-group criteria read from a file (shown as one entry
//   until edited), then "(Blanks)" when some row is blank. OK with every
//   entry checked removes the column's criteria; otherwise it writes a
//   `values` criterion of the checked items.
import { colToLetters, getCell, usedRange, type GridRange } from "../../lib/spreadsheet/address";
import { cellDisplay } from "../../lib/spreadsheet/cellRender";
import { columnValueKind, type FilterCriteria } from "../../lib/spreadsheet/filter";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { plainText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";

export const BLANKS_TEXT = "(Blanks)";
export const DATE_GROUPS_TEXT = "Date groups";

/** "Column B". */
export function columnName(col: number): string {
  return `Column ${colToLetters(col)}`;
}

/** A header cell's display text, or "Column B" when it shows nothing. */
export function headerLabel(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, row: number, col: number, ctx: FormatContext): string {
  const value = getCell(sheet, row, col)?.value ?? null;
  const text = value === null ? "" : plainText(cellDisplay(value, cellStyle(model, effectiveStyleId(sheet, row, col)), sheet, ctx)).trim();
  return text === "" ? columnName(col) : text;
}

/** The two order choices of a column: ascending first. */
export function sortOrderLabels(numeric: boolean): [string, string] {
  return numeric ? ["Smallest to Largest", "Largest to Smallest"] : ["A to Z", "Z to A"];
}

export interface SortColumnChoice {
  col: number; // absolute
  label: string;
  /** Mostly numbers or dates: the order reads Smallest to Largest. */
  numeric: boolean;
}

/** The Sort dialog's column choices over `range` — its columns up to the
 * used range (and `keyCol`) — with or without a header row. */
export function sortColumnChoices(
  model: Pick<SpreadsheetModel, "styles">,
  sheet: Sheet,
  range: GridRange,
  hasHeader: boolean,
  keyCol: number,
  ctx: FormatContext,
): SortColumnChoice[] {
  const used = usedRange(sheet);
  const right = Math.max(range.left, Math.min(range.right, used?.right ?? range.left), Math.min(keyCol, range.right));
  const rows = { top: hasHeader ? range.top + 1 : range.top, bottom: Math.min(range.bottom, used?.bottom ?? range.bottom) };
  const choices: SortColumnChoice[] = [];
  for (let col = range.left; col <= right; col += 1) {
    choices.push({
      col,
      label: hasHeader ? headerLabel(model, sheet, range.top, col, ctx) : columnName(col),
      numeric: columnValueKind(model, sheet, col, rows, ctx) !== "text",
    });
  }
  return choices;
}

/** One checkbox of the popover's list. */
export interface ValueEntry {
  /** "v<i>" for a value, "dates" for the date groups, "blanks". */
  key: string;
  text: string;
  checked: boolean;
}

/** The popover's entries (see the module header). */
export function valueEntries(
  list: { items: readonly { text: string; checked: boolean }[]; blanks: boolean | null },
  criteria: FilterCriteria | undefined,
): ValueEntry[] {
  const entries: ValueEntry[] = list.items.map((item, i) => ({ key: `v${i}`, text: item.text, checked: item.checked }));
  if (criteria?.kind === "values" && criteria.dateGroups.length > 0) {
    entries.push({ key: "dates", text: DATE_GROUPS_TEXT, checked: true });
  }
  if (list.blanks !== null) {
    entries.push({ key: "blanks", text: BLANKS_TEXT, checked: list.blanks });
  }
  return entries;
}

/** What OK writes for the checked keys: `unchanged` when they are the
 * entries' own state, `undefined` (no criteria) when every entry is
 * checked, else a `values` criterion. */
export function valuesCriteria(
  entries: readonly ValueEntry[],
  checked: ReadonlySet<string>,
  criteria: FilterCriteria | undefined,
): FilterCriteria | undefined | "unchanged" {
  if (entries.every((entry) => entry.checked === checked.has(entry.key))) {
    return "unchanged";
  }
  if (entries.every((entry) => checked.has(entry.key))) {
    return undefined;
  }
  return {
    kind: "values",
    values: entries.filter((entry) => entry.key.startsWith("v") && checked.has(entry.key)).map((entry) => entry.text),
    blanks: checked.has("blanks"),
    dateGroups: checked.has("dates") && criteria?.kind === "values" ? criteria.dateGroups : [],
  };
}

/** Narrows a component's runtime-typed entries prop. */
export function asValueEntries(value: unknown): ValueEntry[] {
  return Array.isArray(value) ? (value as ValueEntry[]) : [];
}

/** Narrows a component's runtime-typed column choices prop. */
export function asSortColumnChoices(value: unknown): SortColumnChoice[] {
  return Array.isArray(value) ? (value as SortColumnChoice[]) : [];
}
