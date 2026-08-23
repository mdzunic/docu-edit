import { onBeforeUnmount, onMounted } from "vue";

// SPEC-068 AC-22/23 — a touch long-press equivalent for the desktop right
// click. Touch has no contextmenu gesture, so we detect a still finger held
// past a threshold and hand the caller the point to open the SAME context menu
// the mouse path opens. Critically, a press that MOVES is a text-selection
// drag, not a long-press (AC-23): any movement past a small tolerance cancels
// the timer, so selecting text never trips the menu.

/** The default hold before a still touch becomes a long-press (≈500ms per the
 * spec's touch-equivalents table). */
export const LONG_PRESS_MS = 500;
/** Movement past this many CSS px cancels the long-press: the finger is
 * dragging a selection, not resting. */
export const LONG_PRESS_MOVE_TOLERANCE = 10;

export interface LongPressInfo {
  clientX: number;
  clientY: number;
  target: EventTarget | null;
}

export interface UseLongPressOptions {
  /** The element to watch; re-read on each event so a late-mounted root works. */
  target: () => Element | null;
  /** Called once when a still touch is held past the threshold. */
  onLongPress: (info: LongPressInfo) => void;
  thresholdMs?: number;
  moveTolerance?: number;
}

/** Attaches touch long-press detection to `target`. Only coarse touch pointers
 * arm the timer — a mouse already has a real contextmenu event, and a stylus is
 * out of scope — so this never shadows the desktop right click. */
export function useLongPress(options: UseLongPressOptions) {
  const thresholdMs = options.thresholdMs ?? LONG_PRESS_MS;
  const moveTolerance = options.moveTolerance ?? LONG_PRESS_MOVE_TOLERANCE;

  let timer: number | null = null;
  let startX = 0;
  let startY = 0;
  let fired = false;

  function cancel() {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  function onPointerDown(event: Event) {
    const pointer = event as PointerEvent;
    // Touch only: a mouse/pen already has (or is out of scope for) a real
    // contextmenu gesture. A secondary touch (pinch) also cancels any pending
    // press so a two-finger gesture never opens the menu.
    if (pointer.pointerType !== "touch") {
      return;
    }
    cancel();
    fired = false;
    startX = pointer.clientX;
    startY = pointer.clientY;
    const target = pointer.target;
    timer = window.setTimeout(() => {
      timer = null;
      fired = true;
      options.onLongPress({ clientX: startX, clientY: startY, target });
    }, thresholdMs);
  }

  function onPointerMove(event: Event) {
    if (timer === null) {
      return;
    }
    const pointer = event as PointerEvent;
    if (
      Math.abs(pointer.clientX - startX) > moveTolerance ||
      Math.abs(pointer.clientY - startY) > moveTolerance
    ) {
      // The finger is dragging — a selection, not a long-press (AC-23).
      cancel();
    }
  }

  function onPointerEnd() {
    cancel();
  }

  /** Whether the last completed gesture fired the long-press — call sites use
   * it to swallow the click the browser synthesizes after a touch. */
  function consumedLongPress(): boolean {
    return fired;
  }

  function attach() {
    const el = options.target();
    if (!el) {
      return;
    }
    el.addEventListener("pointerdown", onPointerDown);
    el.addEventListener("pointermove", onPointerMove);
    el.addEventListener("pointerup", onPointerEnd);
    el.addEventListener("pointercancel", onPointerEnd);
  }

  function detach() {
    const el = options.target();
    if (!el) {
      return;
    }
    el.removeEventListener("pointerdown", onPointerDown);
    el.removeEventListener("pointermove", onPointerMove);
    el.removeEventListener("pointerup", onPointerEnd);
    el.removeEventListener("pointercancel", onPointerEnd);
  }

  onMounted(attach);
  onBeforeUnmount(() => {
    cancel();
    detach();
  });

  return { consumedLongPress, cancel };
}

/** Dispatches a synthetic `contextmenu` MouseEvent at the given point so the
 * editor's existing @contextmenu path (spell/table/field hit-testing and all)
 * runs unchanged from a long-press. Targets the element actually under the
 * finger so the hit-tests see the right node. */
export function openContextMenuAtPoint(info: LongPressInfo): void {
  const el =
    (typeof document !== "undefined" && document.elementFromPoint
      ? document.elementFromPoint(info.clientX, info.clientY)
      : null) ??
    (info.target instanceof Element ? info.target : null);
  if (!el) {
    return;
  }
  el.dispatchEvent(
    new MouseEvent("contextmenu", {
      bubbles: true,
      cancelable: true,
      clientX: info.clientX,
      clientY: info.clientY,
    }),
  );
}
