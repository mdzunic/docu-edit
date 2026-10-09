// A sheet's zoom (SPEC-117 "Zoom"), in SPEC-063's vocabulary: the status
// bar's ZoomControl (its fit entries hidden — a sheet has no page to fit),
// the ZOOM_PRESETS ladder for − / + and the View ▸ Zoom menu, Ctrl/Cmd+=/−/0
// and Ctrl/Cmd+wheel. The typed field and every step stay within
// SHEET_ZOOM_MIN–SHEET_ZOOM_MAX (10–400%). Zoom is per sheet and is view
// state: it is read from and written to the sheet's `zoomPercent` through
// useSheetViewState (initially the file's own, e.g. 85% for zoomScale="85").
// It is applied by the grid as a scale transform on its content box, so
// GridMetrics never sees it.
import { computed, onBeforeUnmount, onMounted } from "vue";
import { findShortcut, isOutsideDocumentInput } from "../../lib/shortcuts";
import { clampZoom, stepZoomDown, stepZoomUp } from "../../lib/zoom";
import type { SheetViewState } from "./useSheetViewState";

export const SHEET_ZOOM_MIN = 10;
export const SHEET_ZOOM_MAX = 400;

/** The zoom field's text → a percentage clamped to 10–400, or null for
 * anything that is not a whole number ("85", "85%", " 120 % "). */
export function parseSheetZoom(raw: string): number | null {
  const trimmed = raw.trim().replace(/%$/, "").trim();
  if (!/^\d+$/.test(trimmed)) {
    return null;
  }
  return clampZoom(Number.parseInt(trimmed, 10), SHEET_ZOOM_MIN, SHEET_ZOOM_MAX);
}

export interface SheetZoomDeps {
  viewState: Pick<SheetViewState, "active" | "updateActive" | "activeSheet">;
}

export function useSheetZoom(deps: SheetZoomDeps) {
  /** The sheet on show's zoom in percent (100 when the sheet has none). */
  const percent = computed(() => deps.viewState.active.value.zoomPercent ?? 100);

  function set(value: number): void {
    if (!Number.isFinite(value) || deps.viewState.activeSheet.value?.kind !== "grid") {
      return;
    }
    deps.viewState.updateActive({ zoomPercent: clampZoom(Math.round(value), SHEET_ZOOM_MIN, SHEET_ZOOM_MAX) });
  }

  const zoomIn = (): void => set(stepZoomUp(percent.value));
  const zoomOut = (): void => set(stepZoomDown(percent.value));
  const reset = (): void => set(100);

  /** The status bar field's committed text; unreadable text is ignored (the
   * control reverts to the live value). */
  function applyInput(raw = ""): void {
    const parsed = parseSheetZoom(raw);
    if (parsed !== null) {
      set(parsed);
    }
  }

  /** Ctrl/Cmd+=/−/0 (SPEC-063's view.zoom* shortcuts), anywhere in the view
   * except an input. Returns whether the key was consumed. */
  function onKeydown(event: KeyboardEvent): boolean {
    if (event.defaultPrevented || isOutsideDocumentInput(event.target) || deps.viewState.activeSheet.value?.kind !== "grid") {
      return false;
    }
    const id = findShortcut(event)?.id;
    const action = id === "view.zoomIn" ? zoomIn : id === "view.zoomOut" ? zoomOut : id === "view.zoomReset" ? reset : null;
    if (!action) {
      return false;
    }
    event.preventDefault();
    action();
    return true;
  }

  /** Ctrl/Cmd+wheel over the grid: one ladder step per wheel event. */
  function onWheel(event: WheelEvent): boolean {
    if (!(event.ctrlKey || event.metaKey) || event.deltaY === 0) {
      return false;
    }
    event.preventDefault();
    if (event.deltaY < 0) {
      zoomIn();
    } else {
      zoomOut();
    }
    return true;
  }

  const keydown = (event: KeyboardEvent): void => {
    onKeydown(event);
  };
  onMounted(() => window.addEventListener("keydown", keydown));
  onBeforeUnmount(() => window.removeEventListener("keydown", keydown));

  return { percent, set, zoomIn, zoomOut, reset, applyInput, onKeydown, onWheel };
}

export type SheetZoom = ReturnType<typeof useSheetZoom>;
