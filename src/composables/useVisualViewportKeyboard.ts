import { onBeforeUnmount, onMounted, ref } from "vue";

// SPEC-068 AC-29/30/31 — the on-screen keyboard. When a virtual keyboard opens
// it shrinks the visual viewport under the caret; the app listens to
// visualViewport's resize/scroll and scrolls the DOCUMENT (never resizes the
// layout, which window.innerHeight would invalidate) so the caret rect stays
// inside the still-visible region, and restores the scroll position when the
// keyboard closes. No layout is computed from window.innerHeight at the compact
// tier — that is the value the keyboard invalidates.

/** The vertical adjustment needed to keep the caret above the keyboard: the
 * number of px to add to the scroll container's scrollTop so the caret's bottom
 * sits `pad` px above the visible bottom. 0 when the caret is already visible.
 * Pure so the decision is unit-testable without a real viewport. */
export function keyboardScrollAdjustment(
  caretBottom: number,
  caretTop: number,
  visibleTop: number,
  visibleBottom: number,
  pad = 8,
): number {
  // Caret hidden behind the keyboard (below the visible region): scroll down.
  if (caretBottom > visibleBottom - pad) {
    return Math.ceil(caretBottom - (visibleBottom - pad));
  }
  // Caret above the visible region (rare — a scroll pushed it up): scroll up.
  if (caretTop < visibleTop + pad) {
    return Math.floor(caretTop - (visibleTop + pad));
  }
  return 0;
}

/** True when the visual viewport is meaningfully shorter than the layout
 * viewport — the signal a virtual keyboard (or another inset widget) is open.
 * The 120px floor ignores the browser's own collapsing toolbars. */
export function isKeyboardOpen(layoutHeight: number, visualHeight: number): boolean {
  return layoutHeight - visualHeight > 120;
}

export interface VisualViewportKeyboardDeps {
  /** The scrollable ancestor of the caret (the editor's <main>). */
  getScrollContainer: () => Element | null;
  /** The caret's current client rect, or null when there is no caret. */
  getCaretRect: () => DOMRect | null;
}

export function useVisualViewportKeyboard(deps: VisualViewportKeyboardDeps) {
  /** The keyboard's height in px (0 when closed) — exposed so chrome can lift
   * off the keyboard rather than sitting behind it. */
  const keyboardInset = ref(0);
  let open = false;
  // The document scroll position captured when the keyboard opened, restored on
  // close so nothing is lost (AC-31, AC-44).
  let restoreScrollTop: number | null = null;

  function viewport(): VisualViewport | null {
    return typeof window !== "undefined" ? window.visualViewport : null;
  }

  function onViewportChange() {
    const vv = viewport();
    const container = deps.getScrollContainer();
    if (!vv || !container) {
      return;
    }
    const layoutHeight = window.innerHeight;
    const nowOpen = isKeyboardOpen(layoutHeight, vv.height);
    keyboardInset.value = nowOpen ? Math.max(0, Math.round(layoutHeight - vv.height - vv.offsetTop)) : 0;

    if (nowOpen && !open) {
      // Keyboard just opened: remember where we were so we can restore it.
      restoreScrollTop = container.scrollTop;
    }
    open = nowOpen;

    if (nowOpen) {
      const caret = deps.getCaretRect();
      if (caret) {
        const visibleTop = vv.offsetTop;
        const visibleBottom = vv.offsetTop + vv.height;
        const delta = keyboardScrollAdjustment(caret.bottom, caret.top, visibleTop, visibleBottom);
        if (delta !== 0) {
          container.scrollTop += delta;
        }
      }
    } else if (restoreScrollTop !== null) {
      // Keyboard closed: put the document back where it was (AC-31).
      container.scrollTop = restoreScrollTop;
      restoreScrollTop = null;
    }
  }

  onMounted(() => {
    const vv = viewport();
    if (vv) {
      vv.addEventListener("resize", onViewportChange);
      vv.addEventListener("scroll", onViewportChange);
    }
  });

  onBeforeUnmount(() => {
    const vv = viewport();
    if (vv) {
      vv.removeEventListener("resize", onViewportChange);
      vv.removeEventListener("scroll", onViewportChange);
    }
  });

  return { keyboardInset };
}
