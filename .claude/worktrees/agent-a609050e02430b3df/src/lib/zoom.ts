// The zoom model for the document editor (SPEC-063). Zoom is a
// presentation-layer scale only: the document model, pagination and stored
// geometry stay in centimetres at 96 dpi, and every function here is pure
// arithmetic over integers/pixels so it is unit-testable without a DOM.
//
// Two ladders coexist on purpose:
// - PRESET ladder (50…200): what the − / + buttons and the zoom menu step
//   through (AC-6). The buttons clamp to its ends.
// - INPUT range (25…400): the wider band the editable percentage field
//   accepts (AC-8), clamped to its bounds (AC-9). Fit-width / fit-page land
//   anywhere inside this range.

/** The preset ladder the − / + buttons and the zoom menu offer (AC-4, AC-6). */
export const ZOOM_PRESETS = [50, 75, 90, 100, 125, 150, 200] as const;

/** The − / + buttons never leave the preset ladder's ends (AC-6). */
export const PRESET_MIN: number = ZOOM_PRESETS[0];
export const PRESET_MAX: number = ZOOM_PRESETS[ZOOM_PRESETS.length - 1];

/** The editable percentage field accepts this wider band (AC-8, AC-9). */
export const INPUT_MIN = 25;
export const INPUT_MAX = 400;

/** The default a document with no stored zoom opens at (AC-31). */
export const DEFAULT_ZOOM = 100;

/** Fit width leaves 24 px of breathing room per side (AC-26). */
export const FIT_WIDTH_PADDING = 48;

/** The gutter between pages in the print-layout grid (AC-36). */
export const GRID_GAP = 24;

/** Clamp a percentage into [min, max]. */
export function clampZoom(percent: number, min = INPUT_MIN, max = INPUT_MAX): number {
  if (percent < min) {
    return min;
  }
  if (percent > max) {
    return max;
  }
  return percent;
}

/** The next preset strictly above `percent`, clamped to the ladder's top
 * (AC-6, AC-7). When `percent` sits between two presets, this snaps to the
 * one above; when it is already at or past the top, it stays at PRESET_MAX. */
export function stepZoomUp(percent: number): number {
  for (const preset of ZOOM_PRESETS) {
    if (preset > percent) {
      return preset;
    }
  }
  return PRESET_MAX;
}

/** The next preset strictly below `percent`, clamped to the ladder's bottom
 * (AC-6, AC-7). Mirror of stepZoomUp. */
export function stepZoomDown(percent: number): number {
  for (let i = ZOOM_PRESETS.length - 1; i >= 0; i -= 1) {
    if (ZOOM_PRESETS[i] < percent) {
      return ZOOM_PRESETS[i];
    }
  }
  return PRESET_MIN;
}

/** Parse the editable percentage field (AC-8, AC-9, AC-10). Trims
 * whitespace and an optional trailing "%", requires an integer, and clamps
 * the result into [25, 400]. Returns null for empty or non-integer input so
 * the caller can revert to the last valid zoom on blur (AC-10). */
export function parseZoomInput(raw: string): number | null {
  const trimmed = raw.trim().replace(/%$/, "").trim();
  if (trimmed === "" || !/^-?\d+$/.test(trimmed)) {
    return null;
  }
  const value = Number.parseInt(trimmed, 10);
  if (Number.isNaN(value)) {
    return null;
  }
  return clampZoom(value);
}

/** The zoom percentage at which a page `pageWidthPx` wide (at 100 %) fills
 * the available canvas width minus 48 px, 24 px per side (AC-26). Clamped to
 * the input band so a very narrow pane cannot drive it below 25 %. */
export function fitWidthZoom(availableWidthPx: number, pageWidthPx: number): number {
  if (pageWidthPx <= 0) {
    return DEFAULT_ZOOM;
  }
  const usable = availableWidthPx - FIT_WIDTH_PADDING;
  return clampZoom((usable / pageWidthPx) * 100);
}

/** The zoom percentage at which a page `pageHeightPx` tall (at 100 %) fills
 * the visible canvas height, fitting one whole page (AC-28). */
export function fitPageZoom(visibleHeightPx: number, pageHeightPx: number): number {
  if (pageHeightPx <= 0) {
    return DEFAULT_ZOOM;
  }
  return clampZoom((visibleHeightPx / pageHeightPx) * 100);
}

/** How many page columns fit side by side in the print-layout grid at the
 * current zoom: floor(availableWidth / (pageWidth * zoom + 24)), never below
 * one (AC-36). `zoomFraction` is the scale (1 == 100 %). */
export function gridColumns(availableWidthPx: number, pageWidthPx: number, zoomFraction: number): number {
  const scaledPageWidth = pageWidthPx * zoomFraction;
  const perColumn = scaledPageWidth + GRID_GAP;
  if (perColumn <= 0) {
    return 1;
  }
  return Math.max(1, Math.floor(availableWidthPx / perColumn));
}

/** A rectangle in client (viewport) coordinates — the shape of the parts of
 * DOMRect the coordinate helpers read. */
export interface ClientRect {
  left: number;
  top: number;
}

/** A point in either coordinate space. */
export interface Point {
  x: number;
  y: number;
}

/** Map a client (viewport) point into unscaled content pixels relative to
 * the scaled `.pages` box (AC-21, AC-22). `rect` is the LIVE
 * getBoundingClientRect of `.pages` (already reflecting the CSS scale), so
 * dividing the offset by `zoomFraction` recovers the model-space pixel. */
export function clientToContentPx(clientX: number, clientY: number, rect: ClientRect, zoomFraction: number): Point {
  return {
    x: (clientX - rect.left) / zoomFraction,
    y: (clientY - rect.top) / zoomFraction,
  };
}

/** The inverse of clientToContentPx: place a model-space content pixel back
 * into client coordinates for a position: fixed overlay (AC-22, AC-57).
 * Round-trips to the identity (AC-23). */
export function contentPxToClient(contentX: number, contentY: number, rect: ClientRect, zoomFraction: number): Point {
  return {
    x: rect.left + contentX * zoomFraction,
    y: rect.top + contentY * zoomFraction,
  };
}
