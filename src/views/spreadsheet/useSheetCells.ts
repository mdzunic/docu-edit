// What the sheet paper draws (SPEC-117 "Rendering", "Merges", "Overflow",
// "Border CSS and weights"): for each band of rows the viewport shows — the
// frozen rows, then the scrolled rows — the gridlines, the mounted cells
// (with their resolved CSS, text or rich runs, merge spans and overflow
// boxes) and the resolved border segments, split into the frozen-column and
// scrolled-column regions. Framework-free rules live in
// lib/spreadsheet/cellRender.ts; this composable walks the window and
// assembles plain render items SheetCellLayer binds.
//
// - SPEC-137: cell elements are every position in the window sheetRenderPlan
//   picks — at most GRID_MAX_RENDERED_CELLS of them, the grid's subtree
//   within GRID_MAX_DOM_NODES elements, overscan shrinking first — keyed by
//   slot (`${row mod R}:${col mod C}`, R × C the window's capacity), so a
//   scroll step patches the same elements (settling from a far step it
//   draws the viewport alone, in the same slots: sheetAxis.ts); beyond that the
//   overview canvas paints the sheet (`overviewBands`: stored cells, merges,
//   borders and gridlines of the viewport) and only the active cell keeps an
//   element.
//   Empty cells are transparent, so only a fill hides the gridlines under it.
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
// - SPEC-133: a position inside a table paints its table look (lib's
//   tableCellLook) beneath its own style (styleOverLook — Normal's values
//   counting as unset) and beneath SPEC-129's conditional formats, for the
//   cell elements and the border pass alike. A look is an identity-stable
//   object, so a looked cell's CSS caches under its style id AND its look.
// - SPEC-137: a cell whose box, text or content declarations equal the last
//   plan's under its key keeps the last plan's objects (internItemStyles),
//   so re-rendering an unchanged cell patches no style.
import { computed, type ComputedRef } from "vue";
import { formatA1, type GridRange } from "../../lib/spreadsheet/address";
import {
  cellBoxCss,
  cellCss,
  cellDisplay,
  cellText,
  diagonalLines,
  leftAligned,
  overflowAlignment,
  overflowExtent,
  runCss,
  showFormulasDisplay,
  valueKind,
  type CssDeclarations,
} from "../../lib/spreadsheet/cellRender";
import { conditionalStyle, type CfCellResult } from "../../lib/spreadsheet/conditionalFormats";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import type { MergeIndex } from "../../lib/spreadsheet/merges";
import type { Cell, CellStyle, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { isCellSelected, type SheetSelection } from "../../lib/spreadsheet/selection";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES, type SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext, type FormattedValue } from "../../lib/spreadsheet/numberFormat/render";
import { cellStyle, columnSpanAt, fontMaxDigitWidth } from "../../lib/spreadsheet/styles";
import { lookKey, styleOverLook, tableCellLook } from "../../lib/spreadsheet/tableStyles";
import { DEFAULT_THEME, type DocumentTheme } from "../../lib/core/theme";
import { createBandBuilder, type BuildMode, type GridWindow } from "./sheetBands";
import { internItemStyles, type BandItem, type CellItem } from "./sheetCellItems";
import { cellElementStyle, planRender, type PaperDecorations } from "./sheetRenderPlan";
import { GRID_OVERSCAN_COLS, GRID_OVERSCAN_ROWS, type SheetViewport } from "./useSheetViewport";

export type { BandItem, BorderItem, CellItem, RowItem } from "./sheetCellItems";

/** SPEC-137: the most cell elements the paper ever holds. */
export const GRID_MAX_RENDERED_CELLS = 5_000;
/** SPEC-137: the most elements the whole spreadsheet-grid subtree holds. */
export const GRID_MAX_DOM_NODES = 12_000;


/** SPEC-129: the room a conditional-format icon takes at a cell's left edge
 * — 16 px and a 2 px gap — the text aligning in what remains. */
export const CF_ICON_BOX_PX = 18;

/** Declarations that belong on the text's content span, not its clip box. */
const CONTENT_KEYS = new Set(["transform", "transform-origin", "writing-mode", "text-orientation"]);

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
  /** SPEC-129: a position's conditional-format result; absent = none. */
  conditional?: (row: number, col: number) => CfCellResult | undefined;
  /** SPEC-137: what the paper draws beside each cell, for the node count. */
  decorations?: () => Omit<PaperDecorations, "frozenWidth">;
}

export function useSheetCells(deps: SheetCellsDeps): { bands: ComputedRef<BandItem[]>; overview: ComputedRef<boolean>; window: ComputedRef<GridWindow>; overviewBands: () => BandItem[] } {
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

  /** SPEC-137: what every cell of one build reads — the sheet, the model,
   * the selection and Show formulas — read once per build (buildBands)
   * rather than once per cell, each read a reactive one. */
  let frame: { sheet: Sheet; model: SpreadsheetModel; selection: SheetSelection; showFormulas: boolean } | null = null;
  const sheetNow = (): Sheet => frame?.sheet ?? deps.sheet();
  const modelNow = (): SpreadsheetModel => frame?.model ?? deps.model();

  function styleOf(id: number): CellStyle {
    return cellStyle(modelNow(), id);
  }

  /** SPEC-133: the look of the table covering a position, or undefined. */
  function tableLookAt(row: number, col: number): CellStyle | undefined {
    const sheet = sheetNow();
    const table = sheet.tables?.find(({ ref }) => row >= ref.top && row <= ref.bottom && col >= ref.left && col <= ref.right);
    return table ? tableCellLook(sheet, table, row, col, theme.value) : undefined;
  }

  /** Style `id` over a table look — its own properties win, Normal's
   * (style 0) counting as unset. */
  function overLook(id: number, look: CellStyle | undefined): CellStyle {
    return look ? styleOverLook(look, styleOf(id), styleOf(0)) : styleOf(id);
  }

  function cssFor(id: number, kind: ReturnType<typeof valueKind>, alignHint: FormattedValue["alignHint"], left = false, look?: CellStyle): ReturnType<typeof cellCss> {
    const key = `${id}|${kind}|${alignHint ?? ""}|${left ? "left" : ""}|${look ? lookKey(look) : ""}`;
    let css = cssCache.value.get(key);
    if (!css) {
      const style = overLook(id, look);
      css = cellCss(left ? leftAligned(style) : style, theme.value, kind, alignHint);
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
    return cell?.styleId ?? row?.styleId ?? columnSpanAt(sheetNow().columns, col)?.styleId ?? 0;
  }

  /** One mounted cell (or merge anchor), its element keyed `key`. */
  function cellItem(
    key: string,
    row: Row | undefined,
    cell: Cell | undefined,
    at: { row: number; col: number },
    place: { left: number; top: number; width: number; height: number },
    merge: GridRange | undefined,
  ): CellItem {
    const sheet = sheetNow();
    const styleId = styleIdAt(row, cell, at.col);
    const value = cell?.value ?? null;
    // Showing formulas: a formula as left-aligned text, a value's General text.
    const shown = (frame ? frame.showFormulas : deps.showFormulas?.()) ? showFormulasDisplay(sheet, at.row, at.col, value, modelNow().format) : null;
    // A table look lies beneath the style and a conditional format's dxf
    // (and a scale's fill) over both, before anything reads it — its number
    // format included.
    const look = tableLookAt(at.row, at.col);
    const cf = deps.conditional?.(at.row, at.col);
    const own = overLook(styleId, look);
    const base = shown?.formula ? leftAligned(own) : own;
    const style = cf ? conditionalStyle(base, cf) : base;
    const kind = shown?.formula ? "string" : valueKind(value);
    const mdw = mdwFor(styleId);
    // The value through its number format, fitted to the cell (or merge);
    // a shrink-to-fit cell shrinks its font instead of fitting the text.
    const formatted = shown?.formatted ?? cellDisplay(value, style, sheet, {
      ...formatContext.value,
      ...(style.alignment?.shrinkToFit ? {} : { widthPx: place.width }),
      digitWidthPx: mdw,
    });
    const css = cf ? cellCss(style, theme.value, kind, formatted.alignHint) : cssFor(styleId, kind, formatted.alignHint, shown?.formula, look);
    const textStyle: CssDeclarations = {};
    const contentStyle: CssDeclarations = {};
    for (const [prop, val] of Object.entries(css.text)) {
      (CONTENT_KEYS.has(prop) ? contentStyle : textStyle)[prop] = val;
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
    if (cf?.icon) {
      textLeft += CF_ICON_BOX_PX;
      textWidth -= CF_ICON_BOX_PX;
    }
    textStyle.left = `${textLeft}px`;
    textStyle.width = `${textWidth}px`;
    const runs = cell?.runs && typeof value === "string" && horizontal !== "fill" && !shown?.formula ? cell.runs.map((run) => ({ text: run.text, style: run.font ? runCss(run.font, theme.value) : {} })) : null;
    // A cell with no text box carries its font on its element (sheetRenderPlan).
    const boxStyle = cellElementStyle({ text, runs, fill, cf }, cellBoxCss(css.box, place, merge !== undefined), textStyle);
    return {
      key,
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
      selected: isCellSelected(frame?.selection ?? deps.selection(), at.row, at.col),
      diagonal: diagonalLines(style, theme.value),
      ...(cf ? { cf } : {}),
    };
  }

  const builder = createBandBuilder({ ...deps, cellItem, overLook, tableLookAt, theme: () => theme.value });
  function buildBands(win: GridWindow, mode: BuildMode): BandItem[] {
    frame = { sheet: deps.sheet(), model: deps.model(), selection: deps.selection(), showFormulas: deps.showFormulas?.() === true };
    try {
      return builder.buildBands(win, mode);
    } finally {
      frame = null;
    }
  }
  /** SPEC-137: the last plan's cell items by key (internItemStyles). */
  let lastItems: ReadonlyMap<string, CellItem> = new Map();

  const plan = computed(() => {
    const vp = deps.viewport;
    // Settling from a far scroll step, the grid draws part of the overscan
    // (none at first), its slots still the level's (sheetAxis.ts's header).
    const drawn = vp.overscanDrawn.value;
    const windowAt = (factor: number): GridWindow => {
      const rows = Math.round(GRID_OVERSCAN_ROWS * factor);
      const cols = Math.round(GRID_OVERSCAN_COLS * factor);
      return { rows: vp.rowWindow(Math.round(rows * drawn), rows), cols: vp.colWindow(Math.round(cols * drawn), cols) };
    };
    const deco = { gridLines: true, ...deps.decorations?.(), frozenWidth: vp.frozenW.value };
    const overview = () => ({ window: windowAt(0), bands: buildBands(builder.activeWindow(), "active") });
    const planned = planRender(windowAt, (win) => buildBands(win, "dense"), overview, { cells: GRID_MAX_RENDERED_CELLS, nodes: GRID_MAX_DOM_NODES }, deco);
    lastItems = internItemStyles(planned.bands, lastItems);
    return planned;
  });

  return {
    bands: computed(() => plan.value.bands),
    overview: computed(() => plan.value.overview),
    window: computed(() => plan.value.window),
    /** The overview canvas's paint: the viewport's stored cells, merges,
     * borders and gridlines (read inside its frame effect). */
    overviewBands: () => buildBands({ rows: deps.viewport.rowWindow(0), cols: deps.viewport.colWindow(0) }, "stored"),
  };
}
