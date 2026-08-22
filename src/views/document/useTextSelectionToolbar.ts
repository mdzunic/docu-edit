import { onBeforeUnmount, onMounted, ref, watch } from "vue";

export interface UseTextSelectionToolbarOptions {
  /** The document editing surface; the selection must live inside it. */
  contentRoot: () => HTMLElement | null;
  /** True when the toolbar must never show regardless of the selection: the
   * document is read-only, an object is selected, shape text is being edited,
   * or a track-changes preview is showing (SPEC-061 AC-29). */
  suppressed: () => boolean;
  /** The live zoom fraction (SPEC-063). A zoom change moves the selection in
   * client space without firing selectionchange/scroll/resize, so the toolbar
   * re-measures its anchor whenever this changes to stay registered with the
   * selection under zoom (AC-57). Omit where there is no zoom. */
  zoom?: () => number;
}

/** A selection's anchor point for the contextual toolbar: horizontal centre
 * and the top/bottom of its bounding rect. */
export interface SelectionAnchor {
  left: number;
  top: number;
  bottom: number;
}

/** Any dialog, menu or the find bar currently on screen. Tests each candidate's
 * rendered rects rather than trusting DOM order: the menu bar keeps every menu
 * mounted (display:none) so the first [role=menu] in the DOM is usually closed —
 * getClientRects().length gates on what is actually visible (AC-28). */
function overlayOpen(): boolean {
  const candidates = document.querySelectorAll(
    '[role="dialog"], [role="menu"], [data-testid="find-bar"]',
  );
  for (const candidate of Array.from(candidates)) {
    // getClientRects() ignores display:none (kept-mounted menu-bar menus); the
    // visibility check ignores the toolbar overflow menu, which stays in the
    // DOM as a visibility:hidden position:fixed node while closed.
    if (
      candidate instanceof HTMLElement &&
      candidate.getClientRects().length > 0 &&
      getComputedStyle(candidate).visibility !== "hidden"
    ) {
      return true;
    }
  }
  return false;
}

/** Raises a contextual toolbar over a non-empty range of body text (SPEC-061
 * AC-24..30). Never shown for a collapsed caret (AC-26), while a drag is in
 * progress — it appears on pointer-up (AC-27), while a dialog/menu/find bar is
 * open (AC-28), while an object is selected or the document is otherwise
 * suppressed (AC-29). Escape hides it without touching the selection (AC-30). */
export function useTextSelectionToolbar({ contentRoot, suppressed, zoom }: UseTextSelectionToolbarOptions) {
  const visible = ref(false);
  const anchor = ref<SelectionAnchor>({ left: 0, top: 0, bottom: 0 });

  // A pointer drag is in progress (hide until it ends); and a deliberate
  // Escape dismissal holds the toolbar down until the next fresh interaction.
  let dragging = false;
  let dismissed = false;

  function evaluate() {
    if (dragging || dismissed || suppressed()) {
      visible.value = false;
      return;
    }
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
      visible.value = false;
      return;
    }
    const range = selection.getRangeAt(0);
    const root = contentRoot();
    if (!root || !root.contains(range.commonAncestorContainer)) {
      visible.value = false;
      return;
    }
    if (range.toString().trim() === "") {
      visible.value = false;
      return;
    }
    if (overlayOpen() || typeof range.getBoundingClientRect !== "function") {
      // No geometry to anchor against (jsdom has no Range layout), so there is
      // nothing to show — matches the strip's own no-layout guard.
      visible.value = false;
      return;
    }
    const rect = range.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) {
      visible.value = false;
      return;
    }
    anchor.value = { left: rect.left + rect.width / 2, top: rect.top, bottom: rect.bottom };
    visible.value = true;
  }

  function onSelectionChange() {
    // While dragging, the selection is still forming — wait for pointer-up.
    if (!dragging) {
      evaluate();
    }
  }

  function onPointerDown(event = new PointerEvent("pointerdown")) {
    // A pointerdown on the toolbar's own control is not a new text drag —
    // hiding here would unmount the button before its click lands.
    const target = event.target;
    if (target instanceof Element && target.closest(".contextual-toolbar")) {
      return;
    }
    dragging = true;
    dismissed = false;
    visible.value = false;
  }

  function onPointerUp() {
    dragging = false;
    // Let the selection settle a frame before measuring it.
    requestAnimationFrame(() => evaluate());
  }

  function onScroll() {
    if (visible.value) {
      evaluate();
    }
  }

  /** Escape dismissal — hides the toolbar but leaves the selection intact. */
  function hide() {
    dismissed = true;
    visible.value = false;
  }

  onMounted(() => {
    document.addEventListener("selectionchange", onSelectionChange);
    document.addEventListener("pointerdown", onPointerDown, true);
    document.addEventListener("pointerup", onPointerUp, true);
    document.addEventListener("scroll", onScroll, true);
    window.addEventListener("resize", onScroll);
  });

  onBeforeUnmount(() => {
    document.removeEventListener("selectionchange", onSelectionChange);
    document.removeEventListener("pointerdown", onPointerDown, true);
    document.removeEventListener("pointerup", onPointerUp, true);
    document.removeEventListener("scroll", onScroll, true);
    window.removeEventListener("resize", onScroll);
  });

  // Re-measure the anchor when the zoom changes (SPEC-063 AC-57): the selection
  // moves in client space but fires none of the events above. flush:"post" runs
  // after the scale transform has been applied so the new rect is read.
  if (zoom) {
    watch(zoom, () => {
      if (visible.value) {
        evaluate();
      }
    }, { flush: "post" });
  }

  return { visible, anchor, hide };
}
