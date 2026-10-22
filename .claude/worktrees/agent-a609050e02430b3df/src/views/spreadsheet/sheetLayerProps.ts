// Prop narrowing for the workbook's components (the .vue parser bars casts —
// see AppShell.vue on the espree constraint): each helper coerces a
// runtime-typed prop into the shape the composable that built it declares,
// or derives what the template binds from it.
import { MAX_COLS } from "../../lib/spreadsheet/model";
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

/** SPEC-135: one corner indicator a cell element draws — a note's red
 * triangle or a thread's purple marker (dimmed when resolved). */
export interface AnnotationMarkView {
  kind: "note" | "thread";
  testid: "sheet-note-indicator" | "sheet-thread-indicator";
  resolved: boolean;
}

/** A cell's annotation marks by position (useCellNotes' marksAt). */
export type AnnotationMarksAt = (row: number, col: number) => { note: boolean; thread: "open" | "resolved" | null } | null;

/** The indicators a cell element draws (SPEC-135): none, a note's, a
 * thread's or both. A merge's paint-only copy draws none. */
export function annotationMarksOf(cell: CellItem, at: unknown): AnnotationMarkView[] {
  const marks = typeof at === "function" && !cell.paintOnly ? (at as AnnotationMarksAt)(cell.row, cell.col) : null;
  if (!marks) {
    return [];
  }
  const views: AnnotationMarkView[] = [];
  if (marks.note) {
    views.push({ kind: "note", testid: "sheet-note-indicator", resolved: false });
  }
  if (marks.thread) {
    views.push({ kind: "thread", testid: "sheet-thread-indicator", resolved: marks.thread === "resolved" });
  }
  return views;
}

/** SPEC-130: the cell (by its `data-ref`, a merge's anchor) that an element
 * describes — the active cell and the data-validation input message's id. */
export interface DescribedCell {
  ref: string;
  id: string;
}

/** SheetCellLayer's `describedBy` (useSheetGrid's), or null. */
export function asDescribedCell(value: unknown): DescribedCell | null {
  return value !== null && typeof value === "object" && typeof Reflect.get(value, "ref") === "string" && typeof Reflect.get(value, "id") === "string"
    ? (value as DescribedCell)
    : null;
}

/** SPEC-134: the find bar's marks on the sheet on show (useSheetFind's) —
 * every matching cell, and the current match, by findMarkKey. */
export interface FindMarks {
  sheetId: string;
  matches: ReadonlySet<number>;
  current: number | null;
}

/** A position as a find mark's key. */
export function findMarkKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

/** SheetGrid's and SheetCellLayer's `findMarks`, or null. */
export function asFindMarks(value: unknown): FindMarks | null {
  return value !== null && typeof value === "object" && Reflect.get(value, "matches") instanceof Set ? (value as FindMarks) : null;
}

/** A cell element's identity and ARIA: its id (aria-activedescendant's
 * target), the gridcell role, `sheet-cell` and `data-ref`, its 1-based
 * indices past the header row and column, a merge's spans and the selected
 * state — and, for a cell whose value couldn't be recalculated (SPEC-120
 * "Stale cells"), data-stale="true" and a title saying why; SPEC-130: the
 * active cell's aria-describedby while its input message shows
 * (`describedBy`); SPEC-134: a cell the find bar matches carries
 * data-find-match="true", the current match data-find-current="true" too
 * (`findMarks`). A merge's paint in a further frozen pane carries none of
 * it — only aria-hidden — since the merge's one element is in another pane. */
export function cellAttributes(cell: CellItem, staleTitle?: StaleTitle, describedBy?: DescribedCell | null, findMarks?: FindMarks | null): Record<string, string | number | undefined> {
  if (cell.paintOnly) {
    return { "aria-hidden": "true" };
  }
  const title = staleTitle?.(cell.row, cell.col);
  const key = findMarks ? findMarkKey(cell.row, cell.col) : -1;
  return {
    "data-find-match": findMarks?.matches.has(key) ? "true" : undefined,
    "data-find-current": findMarks && findMarks.current === key ? "true" : undefined,
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
    "aria-describedby": describedBy && describedBy.ref === cell.ref ? describedBy.id : undefined,
  };
}
