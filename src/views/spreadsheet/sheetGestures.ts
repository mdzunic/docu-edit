// The pointer geometry of SPEC-125's two grid gestures, as pure functions
// over the grid's unscaled px (header included) — what useSheetPointer asks
// before it hands a press to the fill handle or to drag-move.
import type { RectItem } from "./useSheetOverlays";
import { COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX } from "./useSheetViewport";

/** How near (screen px) the selection's outline a press starts a drag-move. */
export const MOVE_HIT_PX = 4;

/** Whether the grid point (x, y) lies within MOVE_HIT_PX on screen (at
 * `zoom`) of `outline`, a paper-px rectangle — on either side of its edge. */
export function nearOutline(x: number, y: number, outline: RectItem | null, zoom: number): boolean {
  if (!outline) {
    return false;
  }
  const t = MOVE_HIT_PX / zoom;
  const px = x - ROW_HEADER_WIDTH_PX;
  const py = y - COLUMN_HEADER_HEIGHT_PX;
  const right = outline.left + outline.width;
  const bottom = outline.top + outline.height;
  const outer = px >= outline.left - t && px <= right + t && py >= outline.top - t && py <= bottom + t;
  const inner = px > outline.left + t && px < right - t && py > outline.top + t && py < bottom - t;
  return outer && !inner;
}

/** Whether a client point lies outside the grid's cell area: over the
 * headers or beyond the scroller (`box`, its client rectangle). */
export function outsideCells(box: DOMRect | undefined, clientX: number, clientY: number, zoom: number): boolean {
  if (!box) {
    return false;
  }
  const x = (clientX - box.left) / zoom;
  const y = (clientY - box.top) / zoom;
  return x < ROW_HEADER_WIDTH_PX || y < COLUMN_HEADER_HEIGHT_PX || clientX > box.right || clientY > box.bottom;
}

/** Whether an event's target is the fill handle. */
export function onFillHandle(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest('[data-testid="sheet-fill-handle"]') !== null;
}

/** Whether a press (or hover) on the outline band, with these modifiers, is
 * a drag-move's: not under Shift, nor Cmd on a Mac — those extend or add to
 * the selection (Shift-drag insert is backlog) — while Alt/Option, and Ctrl
 * on Windows/Linux, are the move's own copy modifiers, as in Excel. */
export function outlinePressMoves(event: { shiftKey: boolean; metaKey: boolean }, mac: boolean): boolean {
  return !event.shiftKey && !(mac && event.metaKey);
}
