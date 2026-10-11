// Prop narrowing for the workbook's components (the .vue parser bars casts —
// see AppShell.vue on the espree constraint): each helper coerces a
// runtime-typed prop into the shape the composable that built it declares,
// or derives what the template binds from it.
import type { BandItem, CellItem } from "./sheetCellItems";
import type { FilterButtonView } from "./useSortFilter";

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

/** The header row's filter buttons by position (SPEC-128; useSortFilter's
 * filterButtonAt). */
export type FilterButtonAt = (row: number, col: number) => FilterButtonView | null;

/** The `filter-button` a cell element draws: none or one (a list, so the
 * template binds it once). A merge's paint-only copy draws none. */
export function filterButtonsOf(cell: CellItem, at: unknown): FilterButtonView[] {
  const view = typeof at === "function" && !cell.paintOnly ? (at as FilterButtonAt)(cell.row, cell.col) : null;
  return view ? [view] : [];
}

/** One data bar as SheetCellLayer draws it (SPEC-129): its width (the
 * `data-width-pct` the bar carries) and the bar's and the axis line's
 * styles, percentages of the bar area. */
export interface CfBarView {
  widthPct: number;
  barStyle: Record<string, string>;
  axisStyle?: Record<string, string>;
}

/** A cell's data bar: none or one (a list, so the template binds it once).
 * Solid, or a gradient from the bar colour at its base — the leading edge or
 * the axis — to the colour mixed toward white (computed by the evaluator,
 * no CSS colour function); a 1 px border when it has one; the axis a dashed
 * 1 px line. A merge's paint-only copy draws none. */
export function cfBarsOf(cell: CellItem): CfBarView[] {
  const bar = cell.cf?.dataBar;
  if (!bar || cell.paintOnly) {
    return [];
  }
  const leftward = bar.axisPct !== undefined ? bar.startPct < bar.axisPct : bar.startPct > 0;
  const barStyle: Record<string, string> = {
    left: `${bar.startPct}%`,
    width: `${bar.widthPct}%`,
    background: bar.endColor === undefined ? bar.color : `linear-gradient(to ${leftward ? "left" : "right"}, ${bar.color}, ${bar.endColor})`,
  };
  if (bar.borderColor !== undefined) {
    barStyle.border = `1px solid ${bar.borderColor}`;
  }
  const view: CfBarView = { widthPct: bar.widthPct, barStyle };
  if (bar.axisPct !== undefined) {
    view.axisStyle = { left: `${bar.axisPct}%`, "border-left": `1px dashed ${bar.axisColor ?? "#000000"}` };
  }
  return [view];
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
    "data-cf": cell.cf ? "true" : undefined,
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
