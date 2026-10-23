// The grid's rendering ceiling (SPEC-137 "The grid's ceiling"): which window
// the paper renders as DOM, or whether the overview canvas takes over.
// Framework-free — useSheetCells runs it inside its bands computed.
//
// - Levels: the full overscan (GRID_OVERSCAN_ROWS/COLS), then halved,
//   quartered and none — overscan shrinks to zero before the ceiling is
//   crossed. A level is taken when its window holds at most
//   GRID_MAX_RENDERED_CELLS cell elements and the grid's subtree, counted
//   by paperElementCount plus the headers and GRID_CHROME_ELEMENTS, stays
//   within GRID_MAX_DOM_NODES; a window whose positions alone pass the cell
//   ceiling is skipped without being built.
// - None fits (the viewport alone passes either bound — at 1280 × 720 below
//   ~45% zoom): overview. Only the active cell keeps an element; the canvas
//   paints the rest.
// - paperElementCount mirrors SheetCellLayer's template element for element
//   (an upper bound where a child component's internals vary — an icon's
//   paths), so the bound is counted, not hoped for; SheetGrid.spec.ts holds
//   the DOM to it.
// - A cell that draws no text box (drawsText) carries its text style's font
//   on its own element (cellElementStyle), so a format on an empty cell —
//   SPEC-123's bold, a painted format — still reads from the DOM at no
//   element's cost.
import { fontStack } from "../../lib/core/fontSubstitution";
import type { CssDeclarations } from "../../lib/spreadsheet/cellRender";
import type { BandItem, CellItem } from "./sheetCellItems";
import type { AnnotationMarksAt, FilterButtonAt } from "./sheetLayerProps";
import type { AxisItem } from "./useSheetViewport";

/** The overscan factors tried in order (× GRID_OVERSCAN_ROWS/COLS). */
export const GRID_OVERSCAN_LEVELS: readonly number[] = [1, 0.5, 0.25, 0];

/** What the grid draws beside the paper and its headers — the scroller and
 * its boxes, the overlays, the editor, validation and note layers, a few
 * selection rectangles and drawings — as an allowance in the node budget. */
export const GRID_CHROME_ELEMENTS = 600;

/** The paths an icon (CfIcon, AppIcon) draws at most. */
const ICON_PATHS_MAX = 4;

/** One axis of a render window: its frozen and scrolled items, the visible
 * range and the slot capacity the scrolled items' keys are taken modulo. */
export interface AxisWindow {
  frozen: AxisItem[];
  scrolled: AxisItem[];
  firstVisible: number;
  lastVisible: number;
  capacity: number;
}

/** The template-only decorations SheetCellLayer draws per cell. */
export interface PaperDecorations {
  frozenWidth: number;
  gridLines: boolean;
  filterButton?: FilterButtonAt;
  annotationMarks?: AnnotationMarksAt;
  linkAt?: (row: number, col: number) => boolean;
}

export type RenderPlan<W> = { overview: false; window: W; bands: BandItem[] } | { overview: true; window: W; bands: BandItem[] };

/** Positions with a size along an axis window. */
export function shownItems(axis: AxisWindow): number {
  let n = 0;
  for (const item of axis.frozen) {
    n += item.size > 0 ? 1 : 0;
  }
  for (const item of axis.scrolled) {
    n += item.size > 0 ? 1 : 0;
  }
  return n;
}

/** Whether a cell element draws its text box (an empty cell draws none). */
export function drawsText(cell: Pick<CellItem, "text" | "runs" | "fill" | "cf">): boolean {
  return !cell.cf?.hideValue && (cell.text !== "" || cell.runs !== null || cell.fill !== null);
}

/** The font declarations of a cell's text style. */
const FONT_KEYS = ["font-family", "font-size", "font-weight", "font-style", "text-decoration", "color"] as const;

/** A cell element's style: its box, plus — when it draws no text box — the
 * text style's font, the family wrapped in SPEC-048's stack as
 * SheetCellLayer wraps the text box's (see the module header). */
export function cellElementStyle(cell: Pick<CellItem, "text" | "runs" | "fill" | "cf">, box: CssDeclarations, text: CssDeclarations): CssDeclarations {
  if (drawsText(cell)) {
    return box;
  }
  const style = { ...box };
  for (const key of FONT_KEYS) {
    const value = text[key];
    if (value !== undefined) {
      style[key] = key === "font-family" ? fontStack(value) : value;
    }
  }
  return style;
}

/** The elements one cell element renders (SheetCellLayer's template). */
function cellElements(cell: CellItem, deco: PaperDecorations): number {
  let n = 1;
  const own = !cell.paintOnly;
  if (cell.cf?.dataBar && own) {
    n += cell.cf.dataBar.axisPct !== undefined ? 3 : 2;
  }
  if (cell.cf?.icon && own) {
    n += 2 + ICON_PATHS_MAX;
  }
  if (drawsText(cell)) {
    n += 2;
    if (cell.runs) {
      n += cell.runs.length;
    } else if (cell.fill) {
      n += 3;
    } else if (own && deco.linkAt?.(cell.row, cell.col) === true) {
      n += 1;
    }
  }
  if (cell.diagonal) {
    n += 1 + (cell.diagonal.up ? 1 : 0) + (cell.diagonal.down ? 1 : 0);
  }
  if (own && deco.filterButton?.(cell.row, cell.col)) {
    n += 2 + ICON_PATHS_MAX;
  }
  const marks = own ? deco.annotationMarks?.(cell.row, cell.col) : null;
  if (marks) {
    n += (marks.note ? 1 : 0) + (marks.thread ? 1 : 0);
  }
  return n;
}

/** The paper's elements for `bands`, and how many of them are cells. */
export function paperElementCount(bands: readonly BandItem[], deco: PaperDecorations): { elements: number; cells: number } {
  const frozenCols = deco.frozenWidth > 0;
  let elements = 1;
  let cells = 0;
  for (const band of bands) {
    elements += 1;
    if (deco.gridLines) {
      elements += (frozenCols ? 1 + band.gridRows.length + band.gridColsFrozen.length : 0) + 1 + band.gridRows.length + band.gridColsScrolled.length;
    }
    for (const row of band.rows) {
      elements += 1 + (frozenCols ? 2 : 1);
      for (const cell of row.frozen) {
        elements += cellElements(cell, deco);
      }
      for (const cell of row.scrolled) {
        elements += cellElements(cell, deco);
      }
      cells += row.frozen.length + row.scrolled.length;
    }
    elements += (frozenCols ? 1 + band.bordersFrozen.length : 0) + 1 + band.bordersScrolled.length;
  }
  return { elements, cells };
}

/** Picks the render plan: the widest overscan level whose window fits both
 * ceilings, else the overview (see the module header). */
export function planRender<W extends { rows: AxisWindow; cols: AxisWindow }>(
  windowAt: (factor: number) => W,
  build: (window: W) => BandItem[],
  overview: () => { window: W; bands: BandItem[] },
  limits: { cells: number; nodes: number },
  deco: PaperDecorations,
): RenderPlan<W> {
  for (const factor of GRID_OVERSCAN_LEVELS) {
    const window = windowAt(factor);
    const rows = shownItems(window.rows);
    const cols = shownItems(window.cols);
    if (rows * cols > limits.cells) {
      continue;
    }
    const bands = build(window);
    const count = paperElementCount(bands, deco);
    if (count.cells <= limits.cells && count.elements + rows + cols + GRID_CHROME_ELEMENTS <= limits.nodes) {
      return { overview: false, window, bands };
    }
  }
  return { overview: true, ...overview() };
}
