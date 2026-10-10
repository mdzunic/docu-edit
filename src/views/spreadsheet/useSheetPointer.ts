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
// editor open and the selection where it is). SPEC-124: within
// RESIZE_HIT_PX of a header boundary the cursor is col-resize/row-resize, a
// press drags the boundary (useHeaderResize) instead of selecting, and a
// double-click there autofits (useSheetGrid).
import { onBeforeUnmount, type Ref } from "vue";
import type { CellRef } from "../../lib/spreadsheet/address";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import type { CellEditing } from "./useCellEditing";
import type { HeaderResize } from "./useHeaderResize";
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
  /** SPEC-124: header-boundary resizing; absent = none. */
  resize?: () => HeaderResize | undefined;
}

/** A header-area point's position along its axis, in content px, and the
 * boundary under it — what a resize press, drag or hover reads. */
export interface HeaderPoint {
  axis: "row" | "col";
  pos: number;
  view: number;
  across: number;
}

export function useSheetPointer(deps: SheetPointerDeps) {
  let dragging: "cell" | "row" | "column" | "point" | "resize" | null = null;
  /** The axis of a resize drag (SPEC-124). */
  let resizeAxis: "row" | "col" = "col";
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

  /** A viewport point on an axis (unscaled px, header included) → content
   * px, the frozen band taken into account. */
  function contentAlong(axis: "row" | "col", view: number): number {
    const vp = deps.viewport;
    const offset = view - (axis === "col" ? ROW_HEADER_WIDTH_PX : COLUMN_HEADER_HEIGHT_PX);
    const frozen = axis === "col" ? vp.frozenW.value : vp.frozenH.value;
    return offset < frozen ? offset : offset + (axis === "col" ? vp.sx.value : vp.sy.value);
  }

  /** The header point under (x, y) — a column header gives the column
   * axis, a row header the row axis; null elsewhere (SPEC-124). */
  function headerPoint(x: number, y: number): HeaderPoint | null {
    const hit = hitAt(x, y);
    if (hit.area === "column") {
      return { axis: "col", pos: contentAlong("col", x), view: x - ROW_HEADER_WIDTH_PX, across: y };
    }
    if (hit.area === "row") {
      return { axis: "row", pos: contentAlong("row", y), view: y - COLUMN_HEADER_HEIGHT_PX, across: x };
    }
    return null;
  }

  /** A press on a header boundary starts a resize (SPEC-124). */
  function pressBoundary(x: number, y: number): boolean {
    const resize = deps.resize?.();
    const point = headerPoint(x, y);
    const index = resize && point ? resize.boundaryAt(point.axis, point.pos) : null;
    if (!resize || !point || index === null) {
      return false;
    }
    resize.begin(point.axis, index, point.pos, point.view, point.across);
    return true;
  }

  /** The boundary a double-click landed on, or null (SPEC-124). */
  function boundaryAtClient(clientX: number, clientY: number): { axis: "row" | "col"; index: number } | null {
    const resize = deps.resize?.();
    const at = local(clientX, clientY);
    const point = headerPoint(at.x, at.y);
    const index = resize && point ? resize.boundaryAt(point.axis, point.pos) : null;
    return point && index !== null ? { axis: point.axis, index } : null;
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
    // SPEC-124: a press on a header boundary resizes (never while editing).
    if ((hit.area === "column" || hit.area === "row") && !editing?.editing.value && pressBoundary(point.x, point.y)) {
      dragging = "resize";
      resizeAxis = hit.area === "column" ? "col" : "row";
      (event.currentTarget as Element | null)?.setPointerCapture?.(event.pointerId);
      event.preventDefault();
      return;
    }
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
    if (dragging === "resize") {
      const at = local(event.clientX, event.clientY);
      const view = resizeAxis === "col" ? at.x : at.y;
      deps.resize?.()?.move(contentAlong(resizeAxis, view), view - (resizeAxis === "col" ? ROW_HEADER_WIDTH_PX : COLUMN_HEADER_HEIGHT_PX), resizeAxis === "col" ? at.y : at.x);
      return;
    }
    if (!dragging) {
      // SPEC-124: the resize cursor over a header boundary.
      const resize = deps.resize?.();
      if (resize) {
        const at = local(event.clientX, event.clientY);
        const header = headerPoint(at.x, at.y);
        resize.setHover(header && resize.boundaryAt(header.axis, header.pos) !== null ? header.axis : null);
      }
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
    if (dragging === "resize") {
      dragging = null;
      deps.resize?.()?.finish();
      return;
    }
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

  return { onPointerDown, onPointerMove, onPointerUp, boundaryAtClient };
}
