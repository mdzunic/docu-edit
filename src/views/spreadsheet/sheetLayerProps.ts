// Prop narrowing for the workbook's components (the .vue parser bars casts —
// see AppShell.vue on the espree constraint): each helper coerces a
// runtime-typed prop into the shape the composable that built it declares,
// or derives what the template binds from it.
import type { BandItem, CellItem } from "./useSheetCells";

/** SheetCellLayer's `bands` (useSheetCells' BandItem[]). */
export function asBandItems(value: unknown): BandItem[] {
  return Array.isArray(value) ? (value as BandItem[]) : [];
}

/** A stale cell's title by position on the sheet on show (SPEC-120). */
export type StaleTitle = (row: number, col: number) => string | undefined;

/** SheetCellLayer's `staleTitle` (useSheetGrid's). */
export function asStaleTitle(value: unknown): StaleTitle | undefined {
  return typeof value === "function" ? (value as StaleTitle) : undefined;
}

/** A cell element's identity and ARIA: its id (aria-activedescendant's
 * target), the gridcell role, `sheet-cell` and `data-ref`, its 1-based
 * indices past the header row and column, a merge's spans and the selected
 * state — and, for a cell whose value couldn't be recalculated (SPEC-120
 * "Stale cells"), data-stale="true" and a title saying why. A merge's paint
 * in a further frozen pane carries none of it — only aria-hidden — since the
 * merge's one element is in another pane. */
export function cellAttributes(cell: CellItem, staleTitle?: StaleTitle): Record<string, string | number | undefined> {
  if (cell.paintOnly) {
    return { "aria-hidden": "true" };
  }
  const title = staleTitle?.(cell.row, cell.col);
  return {
    "data-stale": title === undefined ? undefined : "true",
    title,
    id: `sheet-cell-${cell.ref}`,
    role: "gridcell",
    "data-testid": "sheet-cell",
    "data-ref": cell.ref,
    "aria-rowindex": cell.row + 2,
    "aria-colindex": cell.col + 2,
    "aria-rowspan": cell.rowSpan > 1 ? cell.rowSpan : undefined,
    "aria-colspan": cell.colSpan > 1 ? cell.colSpan : undefined,
    "aria-selected": cell.selected ? "true" : "false",
  };
}
