import { nextTick, watch } from "vue";
import type { EditorContext } from "./context";

// Overlay re-measurement on zoom (SPEC-063): a zoom change moves every
// measured overlay just as a scroll does — the `.pages` box changes shape
// between the unscaled and scaled layouts (and its scale changes), so rects
// measured against it (spell underlines, find and comment highlights, the
// vertical ruler's row markers, a selected drawing's handles) go stale until
// the next scroll or selection. This composable re-fires the scroll hooks and
// the drawing-handle measurers once the new layout has rendered, and exposes
// the bundled handle re-measure for other layout-moving moments (undo/redo
// snapshots) to reuse.

interface OverlayRemeasureDeps extends Pick<EditorContext, "zoom"> {
  /** The scroll-hook registry (useFeatureHooks) — fired so every
   * scroll-sensitive overlay re-measures. */
  scrollHooks: Set<() => void>;
  remeasureShapeHandles: () => void;
  remeasureImageHandles: () => void;
  remeasureChartHandles: () => void;
}

export function useOverlayRemeasure(deps: OverlayRemeasureDeps) {
  /** Re-finds the selected drawing's handles across all three drawing kinds —
   * an edit or layout change can move (or remove) any of them. */
  function remeasureDrawingHandles(): void {
    deps.remeasureShapeHandles();
    deps.remeasureImageHandles();
    deps.remeasureChartHandles();
  }

  watch(deps.zoom, () => {
    void nextTick().then(() => {
      deps.scrollHooks.forEach((cb) => cb());
      remeasureDrawingHandles();
    });
  });

  return { remeasureDrawingHandles };
}
