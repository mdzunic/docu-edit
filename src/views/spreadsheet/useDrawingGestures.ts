// One grid's drawing layer and pointer gestures (SPEC-131 "The drawing
// layer"): what SheetGrid paints — the panes and the drawings each mounts
// (sheetDrawingLayout.ts' split over this viewport), with a live preview
// while a drawing is dragged — and the gestures that move, resize and draw.
// The view-level useSheetDrawings owns the selection and runs every command;
// this lends it the grid's focus, visible area and scrolling (`bindGrid`).
//
// - A press on a drawing never reaches the grid's own pointer handling (the
//   cell selection stays as it was): it selects the drawing, captures the
//   pointer on the grid and, once the pointer has travelled
//   DRAW_DRAG_THRESHOLD_CM (3 px), drags a preview; the release is ONE "Move
//   …" step. A press on a handle resizes the same way (resizeKeepsAspect:
//   images and groups keep their aspect on a corner unless Shift is held,
//   shapes only with it) as one "Resize …" step. Pointer deltas are divided by
//   the zoom, so every rectangle stays in 100%-zoom px; a cancelled pointer
//   changes nothing. While a cell is being edited, or a shape is armed, a
//   press on a drawing is the grid's.
// - With a shape armed (Insert ▸ Shapes), a press in the cell area starts the
//   draw gesture: drawGestureBox in cm (core/units.ts converts), Shift
//   constraining, a click placing defaultClickSizeCm at the pointer; the
//   release inserts it. Otherwise a press on a cell while a drawing is
//   selected only deselects it and returns focus to the grid; one on a header
//   deselects and goes on to the grid.
// `wrap` puts these ahead of useSheetPointer's handlers on the grid element.
import { computed, onBeforeUnmount, onMounted, shallowRef, type Ref } from "vue";
import { DRAW_DRAG_THRESHOLD_CM, drawGestureBox, type ShapeGeometry } from "../../lib/core/shapes";
import { cmToPx, pxToCm } from "../../lib/core/units";
import { clampToGrid, type RectPx } from "../../lib/spreadsheet/drawings";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import type { SheetDrawingContent } from "../../lib/spreadsheet/model";
import { layoutDrawingPanes, resizedRect, resizeKeepsAspect, type DrawingHandle, type PaneLayout } from "./sheetDrawingLayout";
import type { DrawingGridHooks, SheetDrawings } from "./useSheetDrawings";
import { COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX, type SheetViewport } from "./useSheetViewport";

export interface DrawingGesturesDeps {
  scroller: Readonly<Ref<HTMLElement | null>>;
  viewport: SheetViewport;
  metrics: () => GridMetrics;
  /** The view's drawings; absent = the grid paints none. */
  drawings: () => SheetDrawings | undefined;
  /** Focus to the grid's focus holder (the in-cell editor). */
  focusGrid: () => void;
}

type Gesture =
  | { kind: "move" | "resize"; id: string; pointerId: number; startX: number; startY: number; rect: RectPx; handle: DrawingHandle; content: SheetDrawingContent; engaged: boolean }
  | { kind: "draw"; pointerId: number; geometry: ShapeGeometry; start: { x: number; y: number } };

/** The grid's own pointer handlers (useSheetPointer's). */
interface GridPointer {
  onPointerDown(event: PointerEvent): void;
  onPointerMove(event: PointerEvent): void;
  onPointerUp(event?: PointerEvent): void;
}

export function useDrawingGestures(deps: DrawingGesturesDeps) {
  let gesture: Gesture | null = null;
  /** A dragged drawing's rectangle until the release (100%-zoom px). */
  const preview = shallowRef<{ id: string; rect: RectPx } | null>(null);
  /** The draw gesture's box (100%-zoom px). */
  const drawBox = shallowRef<RectPx | null>(null);
  const thresholdPx = cmToPx(DRAW_DRAG_THRESHOLD_CM);

  const layout = (): PaneLayout => {
    const vp = deps.viewport;
    return { frozenW: vp.frozenW.value, frozenH: vp.frozenH.value, viewCols: vp.viewCols.value, viewRows: vp.viewRows.value, sx: vp.sx.value, sy: vp.sy.value };
  };

  /** What SheetGrid paints (null: no drawings wired). */
  const layer = computed(() => {
    const drawings = deps.drawings();
    if (!drawings) {
      return null;
    }
    const p = preview.value;
    const shown = p ? drawings.shown.value.map((entry) => (entry.drawing.id === p.id ? { ...entry, rect: p.rect } : entry)) : drawings.shown.value;
    return {
      panes: layoutDrawingPanes(shown, layout(), drawings.selectedId.value),
      drawBox: drawBox.value,
      armed: drawings.armed.value !== null,
      zoom: deps.viewport.z.value,
      theme: drawings.theme.value,
      imageSrc: drawings.imageSrc,
    };
  });

  /** A client point → 100%-zoom sheet px (the frozen bands taken into
   * account), and whether it is in the cell area. */
  function sheetPoint(clientX: number, clientY: number): { x: number; y: number; inCells: boolean } {
    const box = deps.scroller.value?.getBoundingClientRect();
    const vp = deps.viewport;
    const x = (clientX - (box?.left ?? 0)) / vp.z.value - ROW_HEADER_WIDTH_PX;
    const y = (clientY - (box?.top ?? 0)) / vp.z.value - COLUMN_HEADER_HEIGHT_PX;
    return { x: x < vp.frozenW.value ? x : x + vp.sx.value, y: y < vp.frozenH.value ? y : y + vp.sy.value, inCells: x >= 0 && y >= 0 };
  }

  /** Whether a press is on the grid element's own scrollbars. */
  function onScrollbar(event: PointerEvent): boolean {
    const element = deps.scroller.value;
    const box = element?.getBoundingClientRect();
    return Boolean(element && box && element.clientWidth > 0 && (event.clientX - box.left >= element.clientWidth || event.clientY - box.top >= element.clientHeight));
  }

  function capture(event: PointerEvent): void {
    deps.scroller.value?.setPointerCapture?.(event.pointerId);
  }

  /** A press on a drawing (or, with `handle`, on one of its handles). */
  function press(id: string, event: PointerEvent, handle?: DrawingHandle): void {
    const drawings = deps.drawings();
    if (!drawings || drawings.armed.value !== null || drawings.busy()) {
      return;
    }
    event.stopPropagation();
    const entry = drawings.shown.value.find((each) => each.drawing.id === id);
    if (!entry) {
      return;
    }
    if (event.button !== 0) {
      drawings.select(id);
      return;
    }
    event.preventDefault();
    drawings.select(id);
    const kind = handle ? "resize" : "move";
    gesture = { kind, id, pointerId: event.pointerId, startX: event.clientX, startY: event.clientY, rect: entry.rect, handle: handle ?? "se", content: entry.drawing.content, engaged: false };
    capture(event);
  }

  function pressDrawing(id: string, event: PointerEvent): void {
    press(id, event);
  }

  function pressHandle(id: string, handle: DrawingHandle, event: PointerEvent): void {
    press(id, event, handle);
  }

  /** The draw gesture's box with the pointer at `event`. */
  function boxOf(draw: Extract<Gesture, { kind: "draw" }>, event: PointerEvent): RectPx {
    const at = sheetPoint(event.clientX, event.clientY);
    const box = drawGestureBox(
      draw.geometry,
      { xCm: pxToCm(draw.start.x), yCm: pxToCm(draw.start.y) },
      { xCm: pxToCm(Math.max(0, at.x)), yCm: pxToCm(Math.max(0, at.y)) },
      { shiftKey: event.shiftKey },
    );
    return { left: cmToPx(box.leftCm), top: cmToPx(box.topCm), width: cmToPx(box.widthCm), height: cmToPx(box.heightCm) };
  }

  /** A press on the grid itself, ahead of its own handling; true: handled. */
  function gridPointerDown(event: PointerEvent): boolean {
    const drawings = deps.drawings();
    if (!drawings || event.button !== 0 || onScrollbar(event)) {
      return false;
    }
    const at = sheetPoint(event.clientX, event.clientY);
    const geometry = drawings.armed.value;
    if (geometry !== null) {
      if (!at.inCells || drawings.busy()) {
        return false;
      }
      const draw: Gesture = { kind: "draw", pointerId: event.pointerId, geometry, start: { x: Math.max(0, at.x), y: Math.max(0, at.y) } };
      gesture = draw;
      drawBox.value = boxOf(draw, event);
      capture(event);
      event.preventDefault();
      return true;
    }
    if (drawings.selectedId.value === null) {
      return false;
    }
    drawings.deselect(false);
    deps.focusGrid();
    if (at.inCells) {
      event.preventDefault();
      return true;
    }
    return false;
  }

  function onPointerMove(event: PointerEvent): boolean {
    const g = gesture;
    if (!g || g.pointerId !== event.pointerId) {
      return false;
    }
    if (g.kind === "draw") {
      drawBox.value = boxOf(g, event);
      return true;
    }
    if (!g.engaged && Math.max(Math.abs(event.clientX - g.startX), Math.abs(event.clientY - g.startY)) < thresholdPx) {
      return true;
    }
    g.engaged = true;
    const z = deps.viewport.z.value;
    const dx = (event.clientX - g.startX) / z;
    const dy = (event.clientY - g.startY) / z;
    const rect =
      g.kind === "move"
        ? clampToGrid({ ...g.rect, left: g.rect.left + dx, top: g.rect.top + dy })
        : resizedRect(g.rect, g.handle, dx, dy, resizeKeepsAspect(g.content, g.handle, event.shiftKey));
    preview.value = { id: g.id, rect };
    return true;
  }

  function onPointerUp(event?: PointerEvent): boolean {
    const g = gesture;
    if (!g || (event && g.pointerId !== event.pointerId)) {
      return false;
    }
    gesture = null;
    const drawings = deps.drawings();
    const cancelled = !event || event.type === "pointercancel";
    if (g.kind === "draw") {
      const box = event ? boxOf(g, event) : null;
      drawBox.value = null;
      if (!cancelled && box) {
        drawings?.drawShape(g.geometry, box);
      }
      return true;
    }
    const p = preview.value;
    preview.value = null;
    if (!cancelled && g.engaged && p) {
      drawings?.setRect(g.id, p.rect, g.kind === "move" ? "Move" : "Resize");
    }
    return true;
  }

  /** The grid element's pointer handlers: these gestures first. */
  function wrap(pointer: GridPointer): GridPointer {
    return {
      onPointerDown: (event) => {
        if (!gridPointerDown(event)) {
          pointer.onPointerDown(event);
        }
      },
      onPointerMove: (event) => {
        if (!onPointerMove(event)) {
          pointer.onPointerMove(event);
        }
      },
      onPointerUp: (event) => {
        if (!onPointerUp(event)) {
          pointer.onPointerUp(event);
        }
      },
    };
  }

  const hooks: DrawingGridHooks = {
    focusGrid: () => deps.focusGrid(),
    focusDrawing(id) {
      const root = deps.scroller.value;
      const key = id.replace(/["\\]/g, "\\$&");
      const element =
        root?.querySelector<HTMLElement>(`[data-testid="sheet-drawing"][data-drawing-id="${key}"][data-primary="true"]`) ??
        root?.querySelector<HTMLElement>(`[data-testid="sheet-drawing"][data-drawing-id="${key}"]`);
      element?.focus({ preventScroll: true });
    },
    visiblePx: () => ({ width: deps.viewport.viewCols.value, height: deps.viewport.viewRows.value }),
    reveal(rect) {
      const m = deps.metrics();
      const row = m.rowAt(Math.max(0, rect.top));
      const col = m.colAt(Math.max(0, rect.left));
      deps.viewport.ensureVisible({ top: row, left: col, bottom: row, right: col });
    },
  };
  let unbind: (() => void) | null = null;
  onMounted(() => {
    unbind = deps.drawings()?.bindGrid(hooks) ?? null;
  });
  onBeforeUnmount(() => unbind?.());

  return { layer, pressDrawing, pressHandle, wrap, gridPointerDown, onPointerMove, onPointerUp, sheetPoint };
}

export type DrawingGestures = ReturnType<typeof useDrawingGestures>;
