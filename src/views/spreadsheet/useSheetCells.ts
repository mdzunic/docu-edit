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
// - A merge intersecting the window renders once, as one element sized to
//   the range and hosted by the first rendered row it covers (paneMerges),
//   so a hidden or scrolled-off anchor row still shows it; its member cells
//   render nothing. The frozen panes scroll independently, so a merge that
//   crosses a freeze line is that one element in the first pane showing it
//   (its anchor's) plus, in each further pane, a `paintOnly` copy — the
//   same paint, with no id, data-ref, role or ARIA.
// - A string spills into empty neighbours by overflowExtent whenever it has
//   room (its width is never measured, so the text box takes all the room
//   and keeps its alignment); the nearest string stored left (or right) of
//   the window is mounted too, so a long title anchored off-screen still
//   paints into view.
// - SPEC-118: every value renders through its number format (cellDisplay →
//   formatValue) with the cell's (or merge's) width and its font's MDW, in
//   the locale the view hands in: the format's colour overrides the font
//   colour, its alignHint drives General alignment, a `*x` fill becomes a
//   flexible run of its character between the text either side, and a value
//   that does not fit is "####", which never spills. SPEC-119: with Show
//   formulas on, lib's showFormulasDisplay decides the text instead.
import { computed, type ComputedRef } from "vue";
import { findRow, formatA1, getCell, type GridRange } from "../../lib/spreadsheet/address";
import {
  cellBoxCss,
  cellCss,
  cellDisplay,
  cellText,
  diagonalLines,
  edgeElementCss,
  leftAligned,
  offscreenSpillCandidate,
  overflowAlignment,
  overflowExtent,
  resolveEdges,
  runCss,
  showFormulasDisplay,
  valueKind,
  type CellText,
  type CssDeclarations,
  type EdgeSegment,
} from "../../lib/spreadsheet/cellRender";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import { mergeAt, paneMerges, type MergeIndex } from "../../lib/spreadsheet/merges";
import type { Cell, CellStyle, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { isCellSelected, type SheetSelection } from "../../lib/spreadsheet/selection";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES, type SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext, type FormattedValue } from "../../lib/spreadsheet/numberFormat/render";
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
  /** The number locale values render in (useSpreadsheetLocale); absent =
   * the default locale. */
  locale?: () => SpreadsheetLocale | undefined;
  /** View ▸ Show formulas (SPEC-119); absent = values. */
  showFormulas?: () => boolean;
}

/** The border element for one resolved edge, region-relative. */
function borderItem(segment: EdgeSegment, shiftX: number, shiftY: number, index: number): BorderItem {
  return { key: `${segment.orientation}:${segment.x}:${segment.y}:${index}`, style: edgeElementCss(segment, shiftX, shiftY) };
}

export function useSheetCells(deps: SheetCellsDeps): { bands: ComputedRef<BandItem[]> } {
  const theme = computed((): DocumentTheme => deps.model().theme ?? DEFAULT_THEME);
  /** The workbook's format context in the current locale (SPEC-118). */
  const formatContext = computed(() =>
    workbookFormatContext(deps.model(), deps.locale?.() ?? SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]),
  );

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

  function cssFor(id: number, kind: ReturnType<typeof valueKind>, alignHint: FormattedValue["alignHint"], left = false): ReturnType<typeof cellCss> {
    const key = `${id}|${kind}|${alignHint ?? ""}|${left ? "left" : ""}`;
    let css = cssCache.value.get(key);
    if (!css) {
      css = cellCss(left ? leftAligned(styleOf(id)) : styleOf(id), theme.value, kind, alignHint);
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
    const value = cell?.value ?? null;
    // Showing formulas: a formula as left-aligned text, a value's General text.
    const shown = deps.showFormulas?.() ? showFormulasDisplay(sheet, at.row, at.col, value, deps.model().format) : null;
    const style = shown?.formula ? leftAligned(styleOf(styleId)) : styleOf(styleId);
    const kind = shown?.formula ? "string" : valueKind(value);
    const mdw = mdwFor(styleId);
    // The value through its number format, fitted to the cell (or merge);
    // a shrink-to-fit cell shrinks its font instead of fitting the text.
    const formatted = shown?.formatted ?? cellDisplay(value, style, sheet, {
      ...formatContext.value,
      ...(style.alignment?.shrinkToFit ? {} : { widthPx: place.width }),
      digitWidthPx: mdw,
    });
    const css = cssFor(styleId, kind, formatted.alignHint, shown?.formula);
    const textStyle: CssDeclarations = {};
    const contentStyle: CssDeclarations = {};
    for (const [key, val] of Object.entries(css.text)) {
      (CONTENT_KEYS.has(key) ? contentStyle : textStyle)[key] = val;
    }
    if (formatted.color !== undefined) {
      textStyle.color = formatted.color;
    }
    const { text, fill, horizontal, shrunkFontSize } = cellText(formatted, style, kind, place.width, mdw);
    if (shrunkFontSize !== undefined) {
      textStyle["font-size"] = shrunkFontSize;
    }
    let raised = false;
    let textLeft = 0;
    let textWidth = place.width;
    const spill = merge || formatted.overflow === "hash" ? null : overflowAlignment(style, shown?.formula ? formatted.text : value);
    if (spill && text.length > 0) {
      const room = overflowExtent(sheet, at.row, at.col, spill, deps.metrics(), deps.merges(), deps.viewport.frozenCols.value);
      if (room.left > 0 || room.right > 0) {
        raised = true;
        textLeft = -room.left;
        textWidth = place.width + room.left + room.right;
      }
    }
    textStyle.left = `${textLeft}px`;
    textStyle.width = `${textWidth}px`;
    const boxStyle = cellBoxCss(css.box, place, merge !== undefined);
    const runs = cell?.runs && typeof value === "string" && horizontal !== "fill" && !shown?.formula ? cell.runs.map((run) => ({ text: run.text, style: run.font ? runCss(run.font, theme.value) : {} })) : null;
    return {
      key: merge ? `m${merge.top}:${merge.left}` : `${at.row}:${at.col}`,
      ref: formatA1(at),
      row: at.row,
      col: at.col,
      ...place,
      text,
      runs,
      fill,
      boxStyle,
      textStyle,
      contentStyle,
      kindClass: kind === "number" ? "is-number" : kind === "string" ? "is-text" : kind === "empty" ? "is-empty" : "is-centred",
      rowSpan: merge ? merge.bottom - merge.top + 1 : 1,
      colSpan: merge ? merge.right - merge.left + 1 : 1,
      merged: merge !== undefined,
      paintOnly: false,
      raised,
      selected: isCellSelected(deps.selection(), at.row, at.col),
      diagonal: diagonalLines(style, theme.value),
    };
  }

  /** A merge's element hosted by `host` — or, `paintOnly`, its paint in a
   * further pane. */
  function mergeItem(range: GridRange, host: RowItem, colShift: number, paintOnly: boolean): CellItem {
    const sheet = deps.sheet();
    const m = deps.metrics();
    const place = {
      left: m.colLeft(range.left) - colShift,
      top: m.rowTop(range.top) - m.rowTop(host.index),
      width: m.colLeft(range.right + 1) - m.colLeft(range.left),
      height: m.rowTop(range.bottom + 1) - m.rowTop(range.top),
    };
    const anchor = { row: range.top, col: range.left };
    return { ...cellItem(findRow(sheet, range.top), getCell(sheet, range.top, range.left), anchor, place, range), paintOnly };
  }

  /** The unmerged cells of one row in one column window (the merge pass in
   * `bands` draws the merges). */
  function rowCells(row: RowItem, cols: AxisItem[], dense: boolean, colShift: number, scrolled: boolean): CellItem[] {
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
    const mount = (col: number, cell: Cell | undefined): void => {
      if (m.colWidth(col) === 0 || mergeAt(merges, row.index, col) !== undefined) {
        return;
      }
      out.push(cellItem(stored, cell, { row: row.index, col }, { left: m.colLeft(col) - colShift, top: 0, width: m.colWidth(col), height: row.height }, undefined));
    };
    // Off-window strings that may spill into view (never across the freeze
    // line, so only the scrolled region looks).
    if (scrolled) {
      const spill = offscreenSpillCandidate(sheet, row.index, firstCol, lastCol);
      for (const candidate of [spill.left, spill.right]) {
        if (candidate && candidate.col >= deps.viewport.frozenCols.value) {
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
    /** The merges given their one element so far, by anchor — the panes are
     * visited top-left first, so an anchor that shows hosts its element. */
    const placed = new Set<string>();
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
        row.frozen = rowCells(row, cols.frozen, dense, 0, false);
        row.scrolled = rowCells(row, cols.scrolled, dense, scrolledShift, true);
      }
      const lastRow = spec.items.at(-1)!.index;
      const edgeWindow = (items: AxisItem[]): GridRange | null =>
        items.length > 0 ? { top: firstRow, left: items[0]!.index, bottom: lastRow, right: items.at(-1)!.index } : null;
      for (const region of [{ key: "frozen" as const, cols: cols.frozen, shift: 0 }, { key: "scrolled" as const, cols: cols.scrolled, shift: scrolledShift }]) {
        const window = edgeWindow(region.cols);
        for (const { range, host } of window ? paneMerges(deps.merges(), window, rowItems, m.colLeft) : []) {
          const anchor = `${range.top}:${range.left}`;
          host[region.key].push(mergeItem(range, host, region.shift, placed.has(anchor)));
          placed.add(anchor);
        }
      }
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
