import { onBeforeUnmount, onMounted } from "vue";

// SPEC-068 AC-32/33/34 — pinch-to-zoom and two-finger pan for touch. A pinch
// drives SPEC-063's view transform (the CSS scale on the .pages container),
// NOT the browser's page zoom, so the app's own chrome never scales. A
// two-finger pan scrolls the document. We preventDefault on the two-finger
// gesture over the document so it routes to our transform — this is not
// disabling zoom (AC-46): a single-finger gesture and the browser's own pinch
// elsewhere are untouched, and the viewport meta never sets user-scalable=no.

interface Pt {
  clientX: number;
  clientY: number;
}

/** Distance between two touch points. */
export function pinchDistance(a: Pt, b: Pt): number {
  return Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
}

/** Midpoint of two touch points — the pan reference and the zoom focus. */
export function pinchCentroid(a: Pt, b: Pt): { x: number; y: number } {
  return { x: (a.clientX + b.clientX) / 2, y: (a.clientY + b.clientY) / 2 };
}

/** The new zoom percent for a pinch: the starting percent scaled by the ratio
 * of current to initial finger distance. Kept pure for unit testing; the caller
 * clamps to the zoom band. */
export function pinchZoomPercent(initialPercent: number, initialDistance: number, currentDistance: number): number {
  if (initialDistance <= 0) {
    return initialPercent;
  }
  return initialPercent * (currentDistance / initialDistance);
}

export interface UsePinchZoomDeps {
  /** The element the touch listeners attach to (the scroll container). */
  target: () => Element | null;
  /** The scrollable element a two-finger pan moves. */
  getScrollContainer: () => Element | null;
  /** The live zoom percent. */
  getZoomPercent: () => number;
  /** Applies a new zoom percent (the caller clamps to its band). */
  setZoomPercent: (percent: number) => void;
}

export function usePinchZoom(deps: UsePinchZoomDeps) {
  let active = false;
  let startDistance = 0;
  let startZoom = 100;
  let lastCentroid = { x: 0, y: 0 };

  function twoTouches(event: TouchEvent): [Touch, Touch] | null {
    return event.touches.length === 2 ? [event.touches[0], event.touches[1]] : null;
  }

  function onTouchStart(event: Event) {
    const pair = twoTouches(event as TouchEvent);
    if (!pair) {
      active = false;
      return;
    }
    active = true;
    startDistance = pinchDistance(pair[0], pair[1]);
    startZoom = deps.getZoomPercent();
    lastCentroid = pinchCentroid(pair[0], pair[1]);
  }

  function onTouchMove(event: Event) {
    const touchEvent = event as TouchEvent;
    const pair = twoTouches(touchEvent);
    if (!active || !pair) {
      return;
    }
    // Our gesture: keep it off the browser's page zoom / scroll (AC-32).
    touchEvent.preventDefault();

    // Zoom by the change in finger distance.
    const distance = pinchDistance(pair[0], pair[1]);
    deps.setZoomPercent(pinchZoomPercent(startZoom, startDistance, distance));

    // Two-finger pan: scroll the document by the centroid's movement (AC-33).
    const centroid = pinchCentroid(pair[0], pair[1]);
    const container = deps.getScrollContainer();
    if (container) {
      container.scrollLeft -= centroid.x - lastCentroid.x;
      container.scrollTop -= centroid.y - lastCentroid.y;
    }
    lastCentroid = centroid;
  }

  function onTouchEnd(event: Event) {
    if ((event as TouchEvent).touches.length < 2) {
      active = false;
    }
  }

  function attach() {
    const el = deps.target();
    if (!el) {
      return;
    }
    el.addEventListener("touchstart", onTouchStart, { passive: true });
    // Not passive: onTouchMove calls preventDefault to claim the gesture.
    el.addEventListener("touchmove", onTouchMove, { passive: false });
    el.addEventListener("touchend", onTouchEnd);
    el.addEventListener("touchcancel", onTouchEnd);
  }

  function detach() {
    const el = deps.target();
    if (!el) {
      return;
    }
    el.removeEventListener("touchstart", onTouchStart);
    el.removeEventListener("touchmove", onTouchMove);
    el.removeEventListener("touchend", onTouchEnd);
    el.removeEventListener("touchcancel", onTouchEnd);
  }

  onMounted(attach);
  onBeforeUnmount(detach);
}
