import { nextTick, watch } from "vue";
import type { EditorContext } from "./context";

// Overlay re-measurement on layout-moving moments outside the scroll path.
//
// A zoom change (SPEC-063) moves every measured overlay just as a scroll
// does — the `.pages` box changes shape between the unscaled and scaled
// layouts (and its scale changes), so rects measured against it (spell
// underlines, find and comment highlights, the vertical ruler's row
// markers, a selected drawing's handles) go stale until the next scroll or
// selection. A *resize* of the pages container does the same without
// touching the zoom value: opening or closing a side pane (SPEC-065)
// narrows the workspace and the page re-centers inside `.pages`, and a
// window resize re-flows it — SPEC-035's squiggles were drawn displaced
// with the comments pane open because nothing re-fired the measurers on
// that moment. This composable re-fires the scroll hooks and the
// drawing-handle measurers once the new layout has rendered, for both
// triggers, and exposes the bundled handle re-measure for other
// layout-moving moments (undo/redo snapshots) to reuse.
//
// The two triggers never double-fire for one gesture: zoom scales via a CSS
// transform, which changes no layout box, so ResizeObserver stays silent on
// a pure zoom change — and a pane toggle changes no zoom value (except
// under a fit mode, where the extra idle-coalesced pass is harmless).

interface OverlayRemeasureDeps extends Pick<EditorContext, "zoom" | "pageRef" | "onUnmount"> {
  /** The scroll-hook registry (useFeatureHooks) — fired so every
   * scroll-sensitive overlay re-measures. */
  scrollHooks: Set<() => void>;
  remeasureShapeHandles: () => void;
  remeasureImageHandles: () => void;
  remeasureChartHandles: () => void;
  /** Drops the table hover handles (SPEC-022 AC-20/21). Unlike the
   * selection-driven drawing handles, the hover geometry cannot be
   * re-measured here — whether the pointer is still over the table is
   * unknowable after a relayout — and the stale content-frame box otherwise
   * keeps `withinHandleBand` latching the handles at their old position (a
   * fit-page zoom re-tiles the page grid, so the box lands over unrelated
   * content). Clearing lets the next genuine hover re-arm them with fresh
   * geometry. */
  clearTableHover: () => void;
}

export function useOverlayRemeasure(deps: OverlayRemeasureDeps) {
  /** Re-finds the selected drawing's handles across all three drawing kinds —
   * an edit or layout change can move (or remove) any of them. */
  function remeasureDrawingHandles(): void {
    deps.remeasureShapeHandles();
    deps.remeasureImageHandles();
    deps.remeasureChartHandles();
  }

  function remeasureOverlays(): void {
    deps.scrollHooks.forEach((cb) => cb());
    remeasureDrawingHandles();
    deps.clearTableHover();
  }

  watch(deps.zoom, () => {
    void nextTick().then(remeasureOverlays);
  });

  // The pages container's layout box — the origin every overlay rect is
  // measured against (e.g. useSpellCheck's pageRect). ResizeObserver batches
  // its callbacks per frame; the pending flag additionally coalesces a burst
  // (an animated pane, a live window drag) into one re-measure per tick, and
  // each subscriber schedules its own real work (idle/RAF) beyond that.
  let resizePending = false;
  const resizeObserver =
    typeof ResizeObserver === "undefined"
      ? null
      : new ResizeObserver(() => {
          if (resizePending) {
            return;
          }
          resizePending = true;
          void nextTick().then(() => {
            resizePending = false;
            remeasureOverlays();
          });
        });
  watch(
    deps.pageRef,
    (el, previous) => {
      if (!resizeObserver) {
        return;
      }
      if (previous instanceof Element) {
        resizeObserver.unobserve(previous);
      }
      if (el instanceof Element) {
        resizeObserver.observe(el);
      }
    },
    { immediate: true },
  );
  deps.onUnmount(() => resizeObserver?.disconnect());

  return { remeasureDrawingHandles };
}
