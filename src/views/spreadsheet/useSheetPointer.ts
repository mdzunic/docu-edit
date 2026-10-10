// The grid's pointer gestures (SPEC-117 "Pointer"): pointer coordinates are
// converted to unscaled content px (÷ zoom, the frozen bands taken into
// account) and hit-tested with GridMetrics' colAt/rowAt — never by which
// element was under the pointer. A press selects (Shift extends, Mod adds a
// range), a drag extends to the cell under the pointer and auto-scrolls by
// AUTO_SCROLL_PX_PER_FRAME while the pointer is outside the viewport; header
// presses and drags select whole rows or columns; the corner selects all.
// SPEC-122: a press focuses the in-cell editor (the grid's focus holder);
// while a cell is being edited a press goes through useCellEditing first —
// inside the editor it is the browser's (the caret moves), on a cell it
// points (and a drag extends the point) when point mode allows, otherwise
// the edit commits before the selection moves (a formula error keeps the
// editor open and the selection where it is).
import { onBeforeUnmount, type Ref } from "vue";
import type { CellRef } from "../../lib/spreadsheet/address";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import type { CellEditing } from "./useCellEditing";
import { primaryModifier, type SheetSelectionController } from "./useSheetSelection";
import { COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX, type SheetViewport } from "./useSheetViewport";

export const AUTO_SCROLL_PX_PER_FRAME = 16;

/** Where a viewport point (unscaled px from the grid's top-left) lands. */
export type GridHit =
  | { area: "corner" }
  | { area: "column"; col: number }
  | { area: "row"; row: number }
  | { area: "cell"; cell: CellRef };

/** Hit-tests an unscaled viewport point through the metrics. */
export function hitTest(
  x: number,
  y: number,
  metrics: GridMetrics,
  layout: { frozenW: number; frozenH: number; sx: number; sy: number },
): GridHit {
  const px = x - ROW_HEADER_WIDTH_PX;
  const py = y - COLUMN_HEADER_HEIGHT_PX;
  const contentX = px < layout.frozenW ? px : px + layout.sx;
  const contentY = py < layout.frozenH ? py : py + layout.sy;
  if (py < 0 && px < 0) {
    return { area: "corner" };
  }
  if (py < 0) {
    return { area: "column", col: metrics.colAt(Math.max(0, contentX)) };
  }
  if (px < 0) {
    return { area: "row", row: metrics.rowAt(Math.max(0, contentY)) };
  }
  return { area: "cell", cell: { row: metrics.rowAt(contentY), col: metrics.colAt(contentX) } };
}

export interface SheetPointerDeps {
  scroller: Readonly<Ref<HTMLElement | null>>;
  metrics: () => GridMetrics;
  viewport: SheetViewport;
  selection: () => SheetSelectionController;
  /** Cell entry (SPEC-122); absent = a read-only grid. */
  editing?: () => CellEditing | undefined;
  /** Focuses the grid's focus holder (the in-cell editor); default: the
   * scroller. */
  focus?: () => void;
}

export function useSheetPointer(deps: SheetPointerDeps) {
  let dragging: "cell" | "row" | "column" | "point" | null = null;
  let last = { x: 0, y: 0 };
  let frame: number | null = null;

  /** A client point → unscaled px from the grid's top-left, and how far it
   * lies outside the scrolled area (for auto-scroll). */
  function local(clientX: number, clientY: number): { x: number; y: number; dx: number; dy: number } {
    const element = deps.scroller.value;
    const rect = element?.getBoundingClientRect();
    const z = deps.viewport.z.value;
    const x = (clientX - (rect?.left ?? 0)) / z;
    const y = (clientY - (rect?.top ?? 0)) / z;
    const vp = deps.viewport;
    const minX = ROW_HEADER_WIDTH_PX + vp.frozenW.value;
    const minY = COLUMN_HEADER_HEIGHT_PX + vp.frozenH.value;
    const dx = x < minX && vp.sx.value > 0 ? -1 : x > vp.viewW.value ? 1 : 0;
    const dy = y < minY && vp.sy.value > 0 ? -1 : y > vp.viewH.value ? 1 : 0;
    return { x: Math.min(Math.max(x, 0), vp.viewW.value - 1), y: Math.min(Math.max(y, 0), vp.viewH.value - 1), dx, dy };
  }

  function hitAt(x: number, y: number): GridHit {
    const vp = deps.viewport;
    return hitTest(x, y, deps.metrics(), { frozenW: vp.frozenW.value, frozenH: vp.frozenH.value, sx: vp.sx.value, sy: vp.sy.value });
  }

  function extendTo(x: number, y: number): void {
    const hit = hitAt(Math.max(x, ROW_HEADER_WIDTH_PX), Math.max(y, COLUMN_HEADER_HEIGHT_PX));
    if (hit.area === "cell") {
      deps.selection().dragTo(hit.cell);
    }
  }

  function stopAutoScroll(): void {
    if (frame !== null) {
      cancelAnimationFrame(frame);
      frame = null;
    }
  }

  /** While the pointer is past an edge, scroll toward it every frame and keep
   * extending to the cell at the clamped pointer. */
  function autoScroll(): void {
    frame = null;
    if (!dragging) {
      return;
    }
    const element = deps.scroller.value;
    const point = local(last.x, last.y);
    if (!element || (point.dx === 0 && point.dy === 0)) {
      return;
    }
    element.scrollTop += point.dy * AUTO_SCROLL_PX_PER_FRAME;
    element.scrollLeft += point.dx * AUTO_SCROLL_PX_PER_FRAME;
    deps.viewport.onScroll();
    extendTo(point.x, point.y);
    frame = requestAnimationFrame(autoScroll);
  }

  function onPointerDown(event: PointerEvent): void {
    const element = deps.scroller.value;
    const rect = element?.getBoundingClientRect();
    // A press on the scroller's own scrollbars is the browser's.
    if (event.button !== 0 || (element && rect && (event.clientX - rect.left >= element.clientWidth || event.clientY - rect.top >= element.clientHeight) && element.clientWidth > 0)) {
      return;
    }
    const point = local(event.clientX, event.clientY);
    const hit = hitAt(point.x, point.y);
    const mods = { shift: event.shiftKey, mod: primaryModifier(event) };
    const editing = deps.editing?.();
    if (editing) {
      const inEditor = event.target instanceof Element && event.target.closest('[data-testid="cell-editor"]') !== null;
      const outcome = hit.area === "cell" ? editing.pressCell(hit.cell, mods.shift, inEditor) : editing.pressOutside() ? "select" : "handled";
      if (outcome === "native") {
        return;
      }
      if (outcome === "handled") {
        dragging = "point";
        (event.currentTarget as Element | null)?.setPointerCapture?.(event.pointerId);
        event.preventDefault();
        return;
      }
    }
    const selection = deps.selection();
    if (hit.area === "corner") {
      selection.selectEverything();
      return;
    }
    if (hit.area === "column") {
      selection.pressHeader("column", hit.col, mods);
      dragging = "column";
    } else if (hit.area === "row") {
      selection.pressHeader("row", hit.row, mods);
      dragging = "row";
    } else {
      selection.press(hit.cell, mods);
      dragging = "cell";
    }
    last = { x: event.clientX, y: event.clientY };
    if (deps.focus) {
      deps.focus();
    } else {
      deps.scroller.value?.focus({ preventScroll: true });
    }
    (event.currentTarget as Element | null)?.setPointerCapture?.(event.pointerId);
    event.preventDefault();
  }

  function onPointerMove(event: PointerEvent): void {
    if (!dragging) {
      return;
    }
    last = { x: event.clientX, y: event.clientY };
    const point = local(event.clientX, event.clientY);
    if (dragging === "point") {
      const hit = hitAt(Math.max(point.x, ROW_HEADER_WIDTH_PX), Math.max(point.y, COLUMN_HEADER_HEIGHT_PX));
      if (hit.area === "cell") {
        deps.editing?.()?.pointDrag(hit.cell);
      }
      return;
    }
    if (dragging === "cell") {
      extendTo(point.x, point.y);
    } else {
      const hit = hitAt(dragging === "column" ? point.x : 0, dragging === "row" ? point.y : 0);
      if (hit.area === "column") {
        deps.selection().dragTo({ row: 0, col: hit.col });
      } else if (hit.area === "row") {
        deps.selection().dragTo({ row: hit.row, col: 0 });
      }
    }
    if ((point.dx !== 0 || point.dy !== 0) && frame === null) {
      frame = requestAnimationFrame(autoScroll);
    }
  }

  function onPointerUp(): void {
    if (dragging === "point") {
      dragging = null;
      deps.editing?.()?.pointRelease();
      return;
    }
    if (dragging) {
      dragging = null;
      stopAutoScroll();
      deps.selection().release();
    }
  }

  onBeforeUnmount(stopAutoScroll);

  return { onPointerDown, onPointerMove, onPointerUp };
}
