// What the grid draws around and over the sheet paper (SPEC-117 "Frozen panes
// and hidden rows and columns", "Selection", "Dark mode"): the lettered and
// numbered headers of the frozen and scrolled windows (highlighted where the
// selection covers them), the `sheet-hidden-marker` on the header boundary
// where a hidden run sits ("Hidden column G", "Hidden rows 4–6"), the
// `sheet-freeze-line`s, one `sheet-selection` rectangle per range and the
// `sheet-active-cell` outline — and, while a formula is being entered
// (SPEC-122), one `sheet-ref-highlight` per reference on the sheet on show,
// outlined in its `--sheet-ref-N` colour (the pointed one dashed). All of it
// follows the app theme (tokens) — only the paper underneath keeps the
// file's colours.
//
// Positions are unscaled px relative to the paper's origin (the top-left of
// the cell area, past the headers): a frozen row or column sits at its
// metric position, a scrolled one at its metric position minus the scroll
// offset — so a range spanning the freeze is still one contiguous rectangle.
import { computed } from "vue";
import { colToLetters, type GridRange } from "../../lib/spreadsheet/address";
import { hiddenColumnRuns, hiddenRowRuns, type GridMetrics } from "../../lib/spreadsheet/geometry";
import { mergeAt, type MergeIndex } from "../../lib/spreadsheet/merges";
import type { Sheet } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import type { RefHighlight } from "./usePointMode";
import type { AxisItem, SheetViewport } from "./useSheetViewport";

export interface HeaderItem {
  index: number;
  label: string;
  start: number; // relative to its region (frozen: the paper origin; scrolled: the scrolled area)
  size: number;
  selected: boolean;
}

export interface MarkerItem {
  key: string;
  title: string;
  at: number; // paper-relative position of the boundary
}

export interface RectItem {
  key: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SheetOverlaysDeps {
  sheet: () => Sheet;
  metrics: () => GridMetrics;
  merges: () => MergeIndex;
  selection: () => SheetSelection;
  viewport: SheetViewport;
  /** The references of a formula being entered (SPEC-122); absent = none. */
  highlights?: () => RefHighlight[];
}

/** A reference's outline: its rectangle, colour index and dash. */
export interface RefHighlightItem extends RectItem {
  colorIndex: number;
  dashed: boolean;
}

/** The on-screen interval of items start..end (inclusive) along one axis, in
 * paper px: frozen items are fixed, scrolled ones move by `scroll`, and
 * whatever the frozen band covers is not shown. Null when nothing shows. */
export function screenInterval(
  start: number,
  end: number,
  frozenCount: number,
  frozenSize: number,
  scroll: number,
  edge: (i: number) => number,
): [number, number] | null {
  const from = start < frozenCount ? edge(start) : Math.max(frozenSize, edge(start) - scroll);
  const toFrozen = end < frozenCount;
  let to = toFrozen ? edge(end + 1) : edge(end + 1) - scroll;
  if (!toFrozen && start < frozenCount) {
    to = Math.max(to, frozenSize);
  }
  if (!toFrozen && start >= frozenCount && to <= frozenSize) {
    return null;
  }
  return to > from ? [from, to] : null;
}

/** "Hidden column G" / "Hidden columns C–D" / "Hidden row 5" / "Hidden rows 4–6". */
export function hiddenRunTitle(axis: "row" | "column", start: number, end: number): string {
  const name = (i: number): string => (axis === "column" ? colToLetters(i) : String(i + 1));
  if (start === end) {
    return `Hidden ${axis} ${name(start)}`;
  }
  return `Hidden ${axis}s ${name(start)}–${name(end)}`;
}

export function useSheetOverlays(deps: SheetOverlaysDeps) {
  const vp = deps.viewport;

  /** Whether any range covers row/column `index`. */
  function covered(axis: "row" | "column", index: number): boolean {
    return deps.selection().ranges.some((range) =>
      axis === "row" ? range.top <= index && index <= range.bottom : range.left <= index && index <= range.right,
    );
  }

  function headers(axis: "row" | "column", items: AxisItem[]): HeaderItem[] {
    return items
      .filter((item) => item.size > 0)
      .map((item) => ({
        index: item.index,
        label: axis === "column" ? colToLetters(item.index) : String(item.index + 1),
        start: item.start,
        size: item.size,
        selected: covered(axis, item.index),
      }));
  }

  const columnHeaders = computed(() => ({
    frozen: headers("column", vp.cols.value.frozen),
    scrolled: headers("column", vp.cols.value.scrolled),
  }));
  const rowHeaders = computed(() => ({
    frozen: headers("row", vp.rows.value.frozen),
    scrolled: headers("row", vp.rows.value.scrolled),
  }));

  /** Hidden-run markers on the column and row header boundaries. */
  const hiddenMarkers = computed(() => {
    const m = deps.metrics();
    const sheet = deps.sheet();
    const columns: MarkerItem[] = [];
    const rows: MarkerItem[] = [];
    const place = (
      out: MarkerItem[],
      axis: "row" | "column",
      runs: { start: number; end: number }[],
      frozenCount: number,
      frozenSize: number,
      scroll: number,
      edge: (i: number) => number,
    ): void => {
      for (const run of runs) {
        const boundary = edge(run.start);
        const at = run.start < frozenCount ? boundary : boundary - scroll;
        if (run.start >= frozenCount && at < frozenSize) {
          continue;
        }
        out.push({ key: `${axis}:${run.start}`, title: hiddenRunTitle(axis, run.start, run.end), at });
      }
    };
    const cols = vp.cols.value;
    const firstCol = cols.frozen[0]?.index ?? cols.scrolled[0]?.index ?? 0;
    const lastCol = cols.scrolled.at(-1)?.index ?? cols.frozen.at(-1)?.index ?? 0;
    place(columns, "column", hiddenColumnRuns(sheet, firstCol, lastCol), vp.frozenCols.value, vp.frozenW.value, vp.sx.value, m.colLeft);
    const rws = vp.rows.value;
    const firstRow = rws.frozen[0]?.index ?? rws.scrolled[0]?.index ?? 0;
    const lastRow = rws.scrolled.at(-1)?.index ?? rws.frozen.at(-1)?.index ?? 0;
    place(rows, "row", hiddenRowRuns(sheet, firstRow, lastRow), vp.frozenRows.value, vp.frozenH.value, vp.sy.value, m.rowTop);
    return { columns, rows };
  });

  /** A range's on-screen rectangle (paper px), or null when off-screen. */
  function rectOf(range: GridRange, key: string): RectItem | null {
    const m = deps.metrics();
    const x = screenInterval(range.left, range.right, vp.frozenCols.value, vp.frozenW.value, vp.sx.value, m.colLeft);
    const y = screenInterval(range.top, range.bottom, vp.frozenRows.value, vp.frozenH.value, vp.sy.value, m.rowTop);
    if (!x || !y) {
      return null;
    }
    return { key, left: x[0], top: y[0], width: x[1] - x[0], height: y[1] - y[0] };
  }

  /** One rectangle per selected range (a single cell draws only the active
   * outline). */
  const selectionRects = computed((): RectItem[] => {
    const sel = deps.selection();
    const single = sel.ranges.length === 1 && sel.ranges[0]!.top === sel.ranges[0]!.bottom && sel.ranges[0]!.left === sel.ranges[0]!.right;
    if (single) {
      return [];
    }
    return sel.ranges.flatMap((range, i) => {
      const rect = rectOf(range, `range:${i}`);
      return rect ? [rect] : [];
    });
  });

  /** The active cell's outline (its merge's, when it is in one). */
  const activeRect = computed((): RectItem | null => {
    const active = deps.selection().active;
    const merge = mergeAt(deps.merges(), active.row, active.col);
    return rectOf(merge ?? { top: active.row, left: active.col, bottom: active.row, right: active.col }, "active");
  });

  /** The freeze lines: at the frozen columns' right edge and the frozen
   * rows' bottom edge (paper px), when there are any. */
  const freezeLines = computed(() => ({
    vertical: vp.frozenCols.value > 0 ? vp.frozenW.value : null,
    horizontal: vp.frozenRows.value > 0 ? vp.frozenH.value : null,
  }));

  /** The formula's references on this sheet, as outlines (SPEC-122). */
  const refHighlights = computed((): RefHighlightItem[] =>
    (deps.highlights?.() ?? []).flatMap((highlight, i) => {
      const rect = rectOf(highlight.range, `ref:${i}`);
      return rect ? [{ ...rect, colorIndex: highlight.colorIndex, dashed: highlight.dashed }] : [];
    }),
  );

  return { columnHeaders, rowHeaders, hiddenMarkers, selectionRects, activeRect, freezeLines, refHighlights };
}
