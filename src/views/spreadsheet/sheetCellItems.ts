// The render items useSheetCells assembles and SheetCellLayer binds (SPEC-117
// "Rendering"): a band of rows, each row's cells split into the frozen- and
// scrolled-column regions, and the resolved border elements. Plain data —
// no Vue — so the composable that builds them and the helpers that narrow
// them for the component (sheetLayerProps.ts) share one definition.
import type { CellText, CssDeclarations, diagonalLines } from "../../lib/spreadsheet/cellRender";
import type { CfCellResult } from "../../lib/spreadsheet/conditionalFormats";

export interface CellItem {
  key: string;
  ref: string;
  row: number;
  col: number;
  left: number; // relative to the region
  top: number; // relative to the host row
  width: number;
  height: number;
  text: string;
  runs: { text: string; style: CssDeclarations }[] | null;
  /** A number format's `*x` fill (SPEC-118, cellText). */
  fill: CellText["fill"];
  boxStyle: CssDeclarations;
  textStyle: CssDeclarations;
  contentStyle: CssDeclarations;
  kindClass: "is-number" | "is-text" | "is-centred" | "is-empty";
  rowSpan: number;
  colSpan: number;
  merged: boolean;
  /** A merge's paint in a further frozen pane; its element is elsewhere. */
  paintOnly: boolean;
  raised: boolean;
  selected: boolean;
  diagonal: ReturnType<typeof diagonalLines>;
  /** SPEC-129: the cell's conditional-format result — its data bar and
   * icon, and whether it hides its value; absent = none. */
  cf?: CfCellResult;
}

export interface RowItem {
  /** SPEC-137: the row's slot key (its index mod the window's capacity). */
  key: string;
  index: number;
  top: number;
  height: number;
  frozen: CellItem[];
  scrolled: CellItem[];
}

export interface BorderItem {
  key: string;
  style: CssDeclarations;
}

/** A gridline: its slot key and where it sits (SPEC-137 — keyed by slot, so
 * a scroll step moves the same element). */
export interface GridLineItem {
  key: string;
  at: number;
}

export interface BandItem {
  kind: "frozen" | "main";
  top: number;
  height: number;
  rows: RowItem[];
  gridRows: GridLineItem[]; // y of each horizontal gridline, band-relative
  gridColsFrozen: GridLineItem[]; // x of each vertical gridline, region-relative
  gridColsScrolled: GridLineItem[];
  bordersFrozen: BorderItem[];
  bordersScrolled: BorderItem[];
}

/** Whether two declaration objects hold the same declarations. */
function sameDeclarations(a: CssDeclarations, b: CssDeclarations): boolean {
  let count = 0;
  for (const key in a) {
    if (!(key in b) || a[key] !== b[key]) {
      return false;
    }
    count += 1;
  }
  return count === Object.keys(b).length;
}

/** SPEC-137: gives every cell item of `bands` the style objects of `last`'s
 * item under the same key wherever they hold the same declarations, and
 * returns this plan's items by key for the next. Vue's patch compares a
 * style binding by identity before diffing it, so an unchanged cell's
 * re-render writes no style — an edit re-rendered every declaration of the
 * window's ~1,400 cells. */
export function internItemStyles(bands: readonly BandItem[], last: ReadonlyMap<string, CellItem>): Map<string, CellItem> {
  const next = new Map<string, CellItem>();
  const intern = (cell: CellItem, key: string): void => {
    const before = last.get(key);
    if (before) {
      if (sameDeclarations(cell.boxStyle, before.boxStyle)) {
        cell.boxStyle = before.boxStyle;
      }
      if (sameDeclarations(cell.textStyle, before.textStyle)) {
        cell.textStyle = before.textStyle;
      }
      if (sameDeclarations(cell.contentStyle, before.contentStyle)) {
        cell.contentStyle = before.contentStyle;
      }
    }
    next.set(key, cell);
  };
  for (const band of bands) {
    for (const row of band.rows) {
      for (const cell of row.frozen) {
        intern(cell, `${band.kind}|f|${cell.key}`);
      }
      for (const cell of row.scrolled) {
        intern(cell, `${band.kind}|s|${cell.key}`);
      }
    }
  }
  return next;
}
