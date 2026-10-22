// The sheet on show's geometry (SPEC-117 "Geometry", "Merges", "Frozen panes
// and hidden rows and columns"): its GridMetrics, its merge index and its
// frozen counts, each rebuilt only when what it is built from changes — the
// metrics when the sheet's `columns` or `rows` reference (or a default size)
// changes, the merge index when `merges` does — so a view-state change (a
// new sheet object with the same rows) rebuilds nothing. Also the
// NavContext every selection function takes. Metrics are unscaled px: zoom
// is applied outside them.
import { computed, shallowRef, watchEffect } from "vue";
import { buildGridMetrics, type GridMetrics } from "../../lib/spreadsheet/geometry";
import { buildMergeIndex, type MergeIndex } from "../../lib/spreadsheet/merges";
import { MAX_COLS, MAX_ROWS, type Sheet } from "../../lib/spreadsheet/model";
import type { NavContext } from "../../lib/spreadsheet/selection";

/** The neutral metrics/merges an empty view (no sheet) answers with. */
const EMPTY_SHEET: Sheet = {
  id: "",
  name: "",
  visibility: "visible",
  kind: "grid",
  rows: [],
  columns: [],
  defaultColWidthPx: 64,
  defaultRowHeightPt: 15,
  merges: [],
  conditionalFormats: [],
  validations: [],
};

export function useSheetMetrics(sheet: () => Sheet | null) {
  const metrics = shallowRef<GridMetrics>(buildGridMetrics(EMPTY_SHEET));
  const merges = shallowRef<MergeIndex>(buildMergeIndex([]));
  let builtFrom: unknown[] = [];
  let mergesFrom: unknown = null;

  watchEffect(() => {
    const current = sheet() ?? EMPTY_SHEET;
    // A record stored before SPEC-117 may lack the layout fields until the
    // reader-version upgrade replaces it: build with the defaults meanwhile.
    const inputs = [current.columns, current.rows, current.defaultColWidthPx, current.defaultRowHeightPt];
    if (inputs.some((input, i) => input !== builtFrom[i])) {
      builtFrom = inputs;
      metrics.value = buildGridMetrics({ ...EMPTY_SHEET, ...current, columns: current.columns ?? [] });
    }
    if (current.merges !== mergesFrom) {
      mergesFrom = current.merges;
      merges.value = buildMergeIndex(current.merges ?? []);
    }
  });

  /** The frozen rows and columns, clamped to the grid. */
  const frozen = computed(() => {
    const value = sheet()?.frozen;
    return {
      rows: Math.min(Math.max(0, Math.trunc(value?.rows ?? 0)), MAX_ROWS - 1),
      cols: Math.min(Math.max(0, Math.trunc(value?.cols ?? 0)), MAX_COLS - 1),
    };
  });

  const isRowHidden = (r: number): boolean => metrics.value.rowHeight(r) === 0;
  const isColHidden = (c: number): boolean => metrics.value.colWidth(c) === 0;

  /** The NavContext for the sheet on show; `page` is how many rows and
   * columns PageUp/PageDown and Alt+PageUp/PageDown move. */
  function nav(page: { rows: number; cols: number } = { rows: 20, cols: 10 }): NavContext {
    return {
      sheet: sheet() ?? EMPTY_SHEET,
      merges: merges.value,
      isRowHidden,
      isColHidden,
      pageRows: Math.max(1, page.rows),
      pageCols: Math.max(1, page.cols),
    };
  }

  return { metrics, merges, frozen, isRowHidden, isColHidden, nav };
}

export type SheetMetrics = ReturnType<typeof useSheetMetrics>;
