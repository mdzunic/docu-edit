// What the sheet paper draws (SPEC-117 "Rendering", "Merges", "Overflow",
// "Border CSS and weights"): for each band of rows the viewport shows — the
// frozen rows, then the scrolled rows — the gridlines, the mounted cells
// (with their resolved CSS, text or rich runs, merge spans and overflow
// boxes) and the resolved border segments, split into the frozen-column and
// scrolled-column regions. Framework-free rules live in
// lib/spreadsheet/cellRender.ts; this composable walks the window and
// assembles plain render items SheetCellLayer binds.
//
// - Cell elements: every position in the window while it holds at most
//   GRID_MAX_MOUNTED_CELLS positions; beyond that (zoomed far out) only
//   stored cells, merge anchors and the active cell. Empty cells are
//   transparent, so only a fill hides the gridlines under it.
// - A merge intersecting the window renders once, as one element at its
//   anchor sized to the range; its member cells render nothing.
// - A string spills into empty neighbours by overflowExtent; the nearest
//   string stored left (or right) of the window is mounted too, so a long
//   title anchored off-screen still paints into view.
import { computed, type ComputedRef } from "vue";
import { findRow, formatA1, getCell, type GridRange } from "../../lib/spreadsheet/address";
import {
  borderCss,
  cellCss,
  diagonalLines,
  displayText,
  estimateTextWidth,
  fillRepeat,
  offscreenSpillCandidate,
  overflowAlignment,
  overflowExtent,
  resolveEdges,
  resolvedHorizontal,
  runCss,
  shrinkScale,
  valueKind,
  type CssDeclarations,
  type EdgeSegment,
} from "../../lib/spreadsheet/cellRender";
import { CELL_PADDING_PX, type GridMetrics } from "../../lib/spreadsheet/geometry";
import { mergeAt, mergesIn, type MergeIndex } from "../../lib/spreadsheet/merges";
import type { Cell, CellStyle, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { isCellSelected, type SheetSelection } from "../../lib/spreadsheet/selection";
import { cellStyle, columnSpanAt, effectiveStyleId, fontMaxDigitWidth } from "../../lib/spreadsheet/styles";
import { DEFAULT_THEME, type DocumentTheme } from "../../lib/core/theme";
import type { AxisItem, SheetViewport } from "./useSheetViewport";

export const GRID_MAX_MOUNTED_CELLS = 10_000;

/** Declarations that belong on the text's content span, not its clip box. */
const CONTENT_KEYS = new Set(["transform", "transform-origin", "writing-mode", "text-orientation"]);

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
  boxStyle: CssDeclarations;
  textStyle: CssDeclarations;
  contentStyle: CssDeclarations;
  kindClass: "is-number" | "is-text" | "is-centred" | "is-empty";
  rowSpan: number;
  colSpan: number;
  merged: boolean;
  raised: boolean;
  selected: boolean;
  diagonal: ReturnType<typeof diagonalLines>;
}

export interface RowItem {
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

export interface BandItem {
  kind: "frozen" | "main";
  top: number;
  height: number;
  rows: RowItem[];
  gridRows: number[]; // y of each horizontal gridline, band-relative
  gridColsFrozen: number[]; // x of each vertical gridline, region-relative
  gridColsScrolled: number[];
  bordersFrozen: BorderItem[];
  bordersScrolled: BorderItem[];
}

export interface SheetCellsDeps {
  model: () => SpreadsheetModel;
  sheet: () => Sheet;
  metrics: () => GridMetrics;
  merges: () => MergeIndex;
  selection: () => SheetSelection;
  viewport: SheetViewport;
}

/** The border element for one resolved edge, region-relative. */
function borderItem(segment: EdgeSegment, shiftX: number, shiftY: number, index: number): BorderItem {
  const css = borderCss(segment.style);
  const half = Math.ceil(css.width / 2);
  const line = `${css.width}px ${css.lineStyle} ${segment.color}`;
  const style: CssDeclarations =
    segment.orientation === "horizontal"
      ? { left: `${segment.x - shiftX}px`, top: `${segment.y - shiftY - half}px`, width: `${segment.length}px`, "border-top": line }
      : { left: `${segment.x - shiftX - half}px`, top: `${segment.y - shiftY}px`, height: `${segment.length}px`, "border-left": line };
  if (css.opacity !== undefined) {
    style.opacity = String(css.opacity);
  }
  return { key: `${segment.orientation}:${segment.x}:${segment.y}:${index}`, style };
}

export function useSheetCells(deps: SheetCellsDeps): { bands: ComputedRef<BandItem[]> } {
  const theme = computed((): DocumentTheme => deps.model().theme ?? DEFAULT_THEME);

  // Per-style caches, rebuilt with the style table or the theme.
  const cssCache = computed(() => {
    void deps.model().styles;
    void theme.value;
    return new Map<string, ReturnType<typeof cellCss>>();
  });
  const mdwCache = computed(() => {
    void deps.model().styles;
    void theme.value;
    return new Map<number, number>();
  });

  function styleOf(id: number): CellStyle {
    return cellStyle(deps.model(), id);
  }

  function cssFor(id: number, kind: ReturnType<typeof valueKind>): ReturnType<typeof cellCss> {
    const key = `${id}|${kind}`;
    let css = cssCache.value.get(key);
    if (!css) {
      css = cellCss(styleOf(id), theme.value, kind);
      cssCache.value.set(key, css);
    }
    return css;
  }

  /** The maximum digit width of a style's font — for the text estimates. */
  function mdwFor(id: number): number {
    let mdw = mdwCache.value.get(id);
    if (mdw === undefined) {
      mdw = fontMaxDigitWidth(styleOf(id).font, theme.value);
      mdwCache.value.set(id, mdw);
    }
    return mdw;
  }

  /** The style a position shows (effectiveStyleId, with its row and the
   * row's cell already found). */
  function styleIdAt(row: Row | undefined, cell: Cell | undefined, col: number): number {
    return cell?.styleId ?? row?.styleId ?? columnSpanAt(deps.sheet().columns, col)?.styleId ?? 0;
  }

  /** One mounted cell (or merge anchor). */
  function cellItem(
    row: Row | undefined,
    cell: Cell | undefined,
    at: { row: number; col: number },
    place: { left: number; top: number; width: number; height: number },
    merge: GridRange | undefined,
  ): CellItem {
    const sheet = deps.sheet();
    const styleId = styleIdAt(row, cell, at.col);
    const style = styleOf(styleId);
    const value = cell?.value ?? null;
    const kind = valueKind(value);
    const css = cssFor(styleId, kind);
    const textStyle: CssDeclarations = {};
    const contentStyle: CssDeclarations = {};
    for (const [key, val] of Object.entries(css.text)) {
      (CONTENT_KEYS.has(key) ? contentStyle : textStyle)[key] = val;
    }
    let text = displayText(value, sheet);
    const mdw = mdwFor(styleId);
    const horizontal = resolvedHorizontal(style, kind);
    if (horizontal === "fill" && text.length > 0) {
      text = fillRepeat(text, place.width, mdw);
    }
    if (style.alignment?.shrinkToFit && text.length > 0) {
      const scale = shrinkScale(text, Math.max(1, place.width - 2 * CELL_PADDING_PX), mdw);
      if (scale < 1) {
        textStyle["font-size"] = `${((style.font?.sizePt ?? 11) * scale).toFixed(2)}pt`;
      }
    }
    let raised = false;
    let textLeft = 0;
    let textWidth = place.width;
    const spill = merge ? null : overflowAlignment(style, value);
    if (spill && text.length > 0 && estimateTextWidth(text, mdw) + 2 * CELL_PADDING_PX > place.width) {
      const room = overflowExtent(sheet, at.row, at.col, spill, deps.metrics(), deps.merges(), deps.viewport.frozenCols.value);
      if (room.left > 0 || room.right > 0) {
        raised = true;
        textLeft = -room.left;
        textWidth = place.width + room.left + room.right;
      }
    }
    textStyle.left = `${textLeft}px`;
    textStyle.width = `${textWidth}px`;
    const boxStyle: CssDeclarations = {
      ...css.box,
      left: `${place.left}px`,
      top: `${place.top}px`,
      width: `${place.width}px`,
      height: `${place.height}px`,
    };
    if (merge && boxStyle["background-color"] === undefined && boxStyle["background-image"] === undefined) {
      // A merged range hides the gridlines inside it, as Excel draws it.
      boxStyle["background-color"] = "#ffffff";
    }
    const runs = cell?.runs && typeof value === "string" && horizontal !== "fill" ? cell.runs.map((run) => ({ text: run.text, style: run.font ? runCss(run.font, theme.value) : {} })) : null;
    return {
      key: merge ? `m${merge.top}:${merge.left}` : `${at.row}:${at.col}`,
      ref: formatA1(at),
      row: at.row,
      col: at.col,
      ...place,
      text,
      runs,
      boxStyle,
      textStyle,
      contentStyle,
      kindClass: kind === "number" ? "is-number" : kind === "string" ? "is-text" : kind === "empty" ? "is-empty" : "is-centred",
      rowSpan: merge ? merge.bottom - merge.top + 1 : 1,
      colSpan: merge ? merge.right - merge.left + 1 : 1,
      merged: merge !== undefined,
      raised,
      selected: isCellSelected(deps.selection(), at.row, at.col),
      diagonal: diagonalLines(style, theme.value),
    };
  }

  /** The cells of one row in one column window. */
  function rowCells(
    row: RowItem,
    cols: AxisItem[],
    dense: boolean,
    colShift: number,
    rowShift: number,
    firstRow: number,
    scrolled: boolean,
  ): CellItem[] {
    if (cols.length === 0) {
      return [];
    }
    const sheet = deps.sheet();
    const m = deps.metrics();
    const merges = deps.merges();
    const stored = findRow(sheet, row.index);
    const cells = stored?.cells ?? [];
    const firstCol = cols[0]!.index;
    const lastCol = cols.at(-1)!.index;
    const active = deps.selection().active;
    const out: CellItem[] = [];
    const placed = new Set<string>();
    const mount = (col: number, cell: Cell | undefined): void => {
      const merge = mergeAt(merges, row.index, col);
      if (merge) {
        // The anchor renders once: in its own row, or in the band's first
        // row when the anchor row is above the window.
        const host = Math.max(merge.top, firstRow);
        if (host !== row.index || placed.has(`${merge.top}:${merge.left}`)) {
          return;
        }
        placed.add(`${merge.top}:${merge.left}`);
        const anchorCell = merge.top === row.index && merge.left === col ? cell : getCell(sheet, merge.top, merge.left);
        out.push(
          cellItem(findRow(sheet, merge.top), anchorCell, { row: merge.top, col: merge.left }, {
            left: m.colLeft(merge.left) - colShift,
            top: m.rowTop(merge.top) - rowShift - row.top,
            width: m.colLeft(merge.right + 1) - m.colLeft(merge.left),
            height: m.rowTop(merge.bottom + 1) - m.rowTop(merge.top),
          }, merge),
        );
        return;
      }
      if (m.colWidth(col) === 0) {
        return;
      }
      out.push(cellItem(stored, cell, { row: row.index, col }, { left: m.colLeft(col) - colShift, top: 0, width: m.colWidth(col), height: row.height }, undefined));
    };
    // Off-window strings that may spill into view (never across the freeze
    // line, so only the scrolled region looks).
    if (scrolled) {
      const spill = offscreenSpillCandidate(sheet, row.index, firstCol, lastCol);
      for (const candidate of [spill.left, spill.right]) {
        if (candidate) {
          mount(candidate.col, candidate);
        }
      }
    }
    let i = 0;
    while (i < cells.length && cells[i]!.col < firstCol) {
      i += 1;
    }
    if (dense) {
      for (const item of cols) {
        while (i < cells.length && cells[i]!.col < item.index) {
          i += 1;
        }
        mount(item.index, cells[i]?.col === item.index ? cells[i] : undefined);
      }
    } else {
      for (; i < cells.length && cells[i]!.col <= lastCol; i += 1) {
        mount(cells[i]!.col, cells[i]);
      }
      for (const merge of mergesIn(merges, { top: row.index, left: firstCol, bottom: row.index, right: lastCol })) {
        mount(Math.max(merge.left, firstCol), undefined);
      }
      if (active.row === row.index && active.col >= firstCol && active.col <= lastCol && !cells.some((c) => c.col === active.col)) {
        mount(active.col, undefined);
      }
    }
    return out;
  }

  const bands = computed((): BandItem[] => {
    const vp = deps.viewport;
    const m = deps.metrics();
    const sheet = deps.sheet();
    const rows = vp.rows.value;
    const cols = vp.cols.value;
    const positions = (rows.frozen.length + rows.scrolled.length) * (cols.frozen.length + cols.scrolled.length);
    const dense = positions <= GRID_MAX_MOUNTED_CELLS;
    const out: BandItem[] = [];
    const specs = [
      { kind: "frozen" as const, items: rows.frozen, top: 0, height: vp.frozenH.value, rowShift: 0 },
      { kind: "main" as const, items: rows.scrolled, top: vp.frozenH.value, height: Math.max(0, vp.viewRows.value - vp.frozenH.value), rowShift: vp.sy.value + vp.frozenH.value },
    ];
    const scrolledShift = vp.sx.value + vp.frozenW.value;
    for (const spec of specs) {
      if (spec.items.length === 0 || spec.height <= 0) {
        continue;
      }
      const rowItems: RowItem[] = spec.items
        .filter((item) => item.size > 0)
        .map((item) => ({ index: item.index, top: item.start, height: item.size, frozen: [], scrolled: [] }));
      if (rowItems.length === 0) {
        continue;
      }
      const firstRow = rowItems[0]!.index;
      for (const row of rowItems) {
        row.frozen = rowCells(row, cols.frozen, dense, 0, spec.rowShift, firstRow, false);
        row.scrolled = rowCells(row, cols.scrolled, dense, scrolledShift, spec.rowShift, firstRow, true);
      }
      const lastRow = spec.items.at(-1)!.index;
      const edgeWindow = (items: AxisItem[]): GridRange | null =>
        items.length > 0 ? { top: firstRow, left: items[0]!.index, bottom: lastRow, right: items.at(-1)!.index } : null;
      const styleAt = (r: number, c: number): CellStyle => styleOf(effectiveStyleId(sheet, r, c));
      const borders = (items: AxisItem[], shiftX: number): BorderItem[] => {
        const window = edgeWindow(items);
        return window
          ? resolveEdges(window, sheet, styleAt, m, deps.merges(), theme.value).map((segment, i) => borderItem(segment, shiftX, spec.rowShift, i))
          : [];
      };
      out.push({
        kind: spec.kind,
        top: spec.top,
        height: spec.height,
        rows: rowItems,
        gridRows: rowItems.map((row) => row.top + row.height - 1),
        gridColsFrozen: cols.frozen.filter((c) => c.size > 0).map((c) => c.start + c.size - 1),
        gridColsScrolled: cols.scrolled.filter((c) => c.size > 0).map((c) => c.start + c.size - 1),
        bordersFrozen: borders(cols.frozen, 0),
        bordersScrolled: borders(cols.scrolled, scrolledShift),
      });
    }
    return out;
  });

  return { bands };
}
