// The paper's band builder (SPEC-117 "Rendering", SPEC-137 "The grid's
// ceiling"), useSheetCells' walk over a render window: per band of rows (the
// frozen rows, then the scrolled rows) its row items, each row's cells in the
// frozen- and scrolled-column regions, the merges (one element per merge,
// paint-only copies in further panes), the resolved borders and the
// gridlines. Positions and rows are keyed by slot (AxisItem.slot), so a
// scroll step patches the same elements; merges and spill candidates keep
// keys of their own. "dense" builds every position (the DOM paper),
// "stored" only the stored cells (the overview canvas's paint) and "active"
// the active cell alone, with no lines (overview mode's one element). The
// cell items themselves come from useSheetCells' cellItem (the style
// resolution); this module only walks.
import type { GridRange } from "../../lib/spreadsheet/address";
import { findRow, getCell } from "../../lib/spreadsheet/address";
import { edgeElementCss, offscreenSpillCandidate, resolveEdges, type EdgeSegment } from "../../lib/spreadsheet/cellRender";
import { conditionalStyle, type CfCellResult } from "../../lib/spreadsheet/conditionalFormats";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import { mergeAt, paneMerges, type MergeIndex } from "../../lib/spreadsheet/merges";
import type { Cell, CellStyle, Row, Sheet } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import { effectiveStyleId } from "../../lib/spreadsheet/styles";
import type { DocumentTheme } from "../../lib/core/theme";
import type { BandItem, BorderItem, CellItem, RowItem } from "./sheetCellItems";
import type { AxisWindow } from "./sheetRenderPlan";
import type { AxisItem, SheetViewport } from "./useSheetViewport";

/** A render window: one AxisWindow per axis. */
export type GridWindow = { rows: AxisWindow; cols: AxisWindow };
/** What a pass builds (see the module header). */
export type BuildMode = "dense" | "stored" | "active";

export interface BandBuilderDeps {
  sheet: () => Sheet;
  metrics: () => GridMetrics;
  merges: () => MergeIndex;
  selection: () => SheetSelection;
  viewport: SheetViewport;
  conditional?: (row: number, col: number) => CfCellResult | undefined;
  theme: () => DocumentTheme;
  cellItem: (
    key: string,
    row: Row | undefined,
    cell: Cell | undefined,
    at: { row: number; col: number },
    place: { left: number; top: number; width: number; height: number },
    merge: GridRange | undefined,
  ) => CellItem;
  overLook: (id: number, look: CellStyle | undefined) => CellStyle;
  tableLookAt: (row: number, col: number) => CellStyle | undefined;
}

/** The border element for one resolved edge, region-relative — keyed by its
 * place on the sheet, stable as the window slides. */
function borderItem(segment: EdgeSegment, shiftX: number, shiftY: number): BorderItem {
  return { key: `${segment.orientation}:${segment.x}:${segment.y}:${segment.length}`, style: edgeElementCss(segment, shiftX, shiftY) };
}

/** A one-item axis window (the active cell's row or column). */
function singleAxis(index: number, frozen: boolean, start: number, size: number): AxisWindow {
  const item = { index, start, size, slot: 0 };
  return { frozen: frozen ? [item] : [], scrolled: frozen ? [] : [item], firstVisible: index, lastVisible: index, capacity: 1 };
}

export function createBandBuilder(c: BandBuilderDeps) {
  /** A merge's element hosted by `host` — or, `paintOnly`, its paint in a
   * further pane. */
  function mergeItem(range: GridRange, host: RowItem, colShift: number, paintOnly: boolean): CellItem {
    const sheet = c.sheet();
    const m = c.metrics();
    const place = {
      left: m.colLeft(range.left) - colShift,
      top: m.rowTop(range.top) - m.rowTop(host.index),
      width: m.colLeft(range.right + 1) - m.colLeft(range.left),
      height: m.rowTop(range.bottom + 1) - m.rowTop(range.top),
    };
    const anchor = { row: range.top, col: range.left };
    return { ...c.cellItem(`m${range.top}:${range.left}`, findRow(sheet, range.top), getCell(sheet, range.top, range.left), anchor, place, range), paintOnly };
  }

  /** The unmerged cells of one row in one column window (the merge pass in
   * `bands` draws the merges): every position (dense, keyed by slot) or the
   * stored cells; off-window strings that may spill into view in the
   * scrolled region (never in "active"). */
  function rowCells(row: RowItem, rowSlot: number, cols: AxisItem[], mode: BuildMode, colShift: number, scrolled: boolean): CellItem[] {
    if (cols.length === 0) {
      return [];
    }
    const sheet = c.sheet();
    const m = c.metrics();
    const merges = c.merges();
    const stored = findRow(sheet, row.index);
    const cells = stored?.cells ?? [];
    const firstCol = cols[0]!.index;
    const lastCol = cols.at(-1)!.index;
    const out: CellItem[] = [];
    const mount = (key: string, col: number, cell: Cell | undefined): void => {
      if (m.colWidth(col) === 0 || mergeAt(merges, row.index, col) !== undefined) {
        return;
      }
      out.push(c.cellItem(key, stored, cell, { row: row.index, col }, { left: m.colLeft(col) - colShift, top: 0, width: m.colWidth(col), height: row.height }, undefined));
    };
    // Off-window strings that may spill into view (never across the freeze
    // line, so only the scrolled region looks); keys of their own.
    if (scrolled && mode !== "active") {
      const spill = offscreenSpillCandidate(sheet, row.index, firstCol, lastCol);
      if (spill.left && spill.left.col >= c.viewport.frozenCols.value) {
        mount(`s${rowSlot}:l`, spill.left.col, spill.left);
      }
      if (spill.right && spill.right.col >= c.viewport.frozenCols.value) {
        mount(`s${rowSlot}:r`, spill.right.col, spill.right);
      }
    }
    let i = 0;
    while (i < cells.length && cells[i]!.col < firstCol) {
      i += 1;
    }
    if (mode !== "stored") {
      for (const item of cols) {
        while (i < cells.length && cells[i]!.col < item.index) {
          i += 1;
        }
        mount(`${rowSlot}:${item.slot}`, item.index, cells[i]?.col === item.index ? cells[i] : undefined);
      }
    } else {
      for (; i < cells.length && cells[i]!.col <= lastCol; i += 1) {
        mount(`${row.index}:${cells[i]!.col}`, cells[i]!.col, cells[i]);
      }
    }
    return out;
  }

  /** The bands of one window (SheetCellLayer's items, or the overview's). */
  function buildBands(win: GridWindow, mode: BuildMode): BandItem[] {
    const vp = c.viewport;
    const m = c.metrics();
    const sheet = c.sheet();
    const { rows, cols } = win;
    const lines = mode !== "active";
    const out: BandItem[] = [];
    const specs = [
      { kind: "frozen" as const, items: rows.frozen, top: 0, height: vp.frozenH.value, rowShift: 0 },
      { kind: "main" as const, items: rows.scrolled, top: vp.frozenH.value, height: Math.max(0, vp.viewRows.value - vp.frozenH.value), rowShift: vp.sy.value + vp.frozenH.value },
    ];
    const scrolledShift = vp.sx.value + vp.frozenW.value;
    /** The merges given their one element so far, by anchor — the panes are
     * visited top-left first, so an anchor that shows hosts its element. */
    const placed = new Set<string>();
    for (const spec of specs) {
      if (spec.items.length === 0 || spec.height <= 0) {
        continue;
      }
      const shown = spec.items.filter((item) => item.size > 0);
      const rowItems: RowItem[] = shown.map((item) => ({ key: `${item.slot}`, index: item.index, top: item.start, height: item.size, frozen: [], scrolled: [] }));
      if (rowItems.length === 0) {
        continue;
      }
      const firstRow = rowItems[0]!.index;
      rowItems.forEach((row, n) => {
        row.frozen = rowCells(row, shown[n]!.slot, cols.frozen, mode, 0, false);
        row.scrolled = rowCells(row, shown[n]!.slot, cols.scrolled, mode, scrolledShift, true);
      });
      const lastRow = spec.items.at(-1)!.index;
      const edgeWindow = (items: AxisItem[]): GridRange | null =>
        items.length > 0 ? { top: firstRow, left: items[0]!.index, bottom: lastRow, right: items.at(-1)!.index } : null;
      for (const region of [{ key: "frozen" as const, cols: cols.frozen, shift: 0 }, { key: "scrolled" as const, cols: cols.scrolled, shift: scrolledShift }]) {
        const window = edgeWindow(region.cols);
        for (const { range, host } of window ? paneMerges(c.merges(), window, rowItems, m.colLeft) : []) {
          const anchor = `${range.top}:${range.left}`;
          host[region.key].push(mergeItem(range, host, region.shift, placed.has(anchor)));
          placed.add(anchor);
        }
      }
      const styleAt = (r: number, k: number): CellStyle => {
        const base = c.overLook(effectiveStyleId(sheet, r, k), c.tableLookAt(r, k));
        const cf = c.conditional?.(r, k);
        return cf ? conditionalStyle(base, cf) : base;
      };
      const borders = (items: AxisItem[], shiftX: number): BorderItem[] => {
        const window = lines ? edgeWindow(items) : null;
        return window ? resolveEdges(window, sheet, styleAt, m, c.merges(), c.theme()).map((segment) => borderItem(segment, shiftX, spec.rowShift)) : [];
      };
      const gridCols = (items: AxisItem[]) => (lines ? items.filter((item) => item.size > 0).map((item) => ({ key: `v${item.slot}`, at: item.start + item.size - 1 })) : []);
      out.push({
        kind: spec.kind,
        top: spec.top,
        height: spec.height,
        rows: rowItems,
        gridRows: lines ? rowItems.map((row) => ({ key: `h${row.key}`, at: row.top + row.height - 1 })) : [],
        gridColsFrozen: gridCols(cols.frozen),
        gridColsScrolled: gridCols(cols.scrolled),
        bordersFrozen: borders(cols.frozen, 0),
        bordersScrolled: borders(cols.scrolled, scrolledShift),
      });
    }
    return out;
  }

  /** The window of the active cell alone (overview mode's one element). */
  function activeWindow(): GridWindow {
    const vp = c.viewport;
    const m = c.metrics();
    const { row, col } = c.selection().active;
    const rowFrozen = row < vp.frozenRows.value;
    const colFrozen = col < vp.frozenCols.value;
    return {
      rows: singleAxis(row, rowFrozen, m.rowTop(row) - (rowFrozen ? 0 : vp.sy.value + vp.frozenH.value), m.rowHeight(row)),
      cols: singleAxis(col, colFrozen, m.colLeft(col) - (colFrozen ? 0 : vp.sx.value + vp.frozenW.value), m.colWidth(col)),
    };
  }

  return { buildBands, activeWindow };
}
