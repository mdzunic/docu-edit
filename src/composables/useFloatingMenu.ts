import { getCurrentInstance, onBeforeUnmount, ref, watch } from "vue";

export interface UseFloatingMenuOptions {
  /** True while the popup is open; the popup only renders then, and it is
   * (re)positioned at the moment it opens. */
  open: () => boolean;
  /** The button the popup anchors to; measured via getBoundingClientRect. */
  anchor: () => Element | null;
  /** The popup itself; its width is measured so it can be kept on-screen. */
  menu: () => Element | null;
  /** Gap between the anchor and the popup, in px. */
  offsetY?: number;
  /** Minimum distance to keep to the viewport's left/right edges, in px. */
  margin?: number;
  /** Horizontal placement: flush with the anchor's left edge (the default,
   * how a dropdown menu hangs off its button) or centred under it (how a
   * tooltip points at its control). */
  align?: "start" | "center";
  /** Placement: below the anchor (the default, for toolbar buttons near the
   * top), above it (for a status-bar control anchored at the bottom of the
   * viewport, where "below" would open off-screen — flipped below when there
   * is no room above), or beside it ("right" —
   * a submenu flyout, top-aligned with its trigger and flipped to the left
   * side when the right edge would run off-screen). */
  placement?: "below" | "above" | "right";
}

/** Anchors a position:fixed popup below (or beside) its trigger. The toolbar that hosts
 * these pickers scrolls horizontally (overflow-x: auto) and clips any content
 * taller than one row, so an ordinary absolutely-positioned menu is cut off.
 * A fixed popup escapes that clipping; on scroll/resize we re-measure the
 * anchor so it stays pinned to its button even while the strip scrolls. */
export function useFloatingMenu({
  open,
  anchor,
  menu,
  offsetY = 4,
  margin = 8,
  align = "start",
  placement = "below",
}: UseFloatingMenuOptions) {
  // `position: fixed` from the very first render, before any measurement:
  // the popup must never take part in the toolbar's flow, or opening it
  // would grow the strip and push the anchor down — and the measurement
  // below, which runs on that perturbed layout, would then place the popup
  // far beneath its button. Only left/top wait for the measurement.
  const menuStyle = ref<Record<string, string>>({ position: "fixed" });

  function position() {
    if (!open()) {
      return;
    }
    const anchorEl = anchor();
    if (!anchorEl) {
      return;
    }
    const anchorRect = anchorEl.getBoundingClientRect();
    const menuEl = menu();
    const menuRect = menuEl ? menuEl.getBoundingClientRect() : null;
    const menuWidth = menuRect ? menuRect.width : 0;
    const maxLeft = window.innerWidth - menuWidth - margin;
    if (placement === "right") {
      // Beside the trigger: flush with its right edge, top-aligned. When the
      // flyout would run off the right of the viewport, flip it to the left
      // side of the trigger (the classic submenu flip), then clamp.
      const rightward = anchorRect.right + offsetY;
      const wantedSide =
        rightward > maxLeft ? anchorRect.left - offsetY - menuWidth : rightward;
      const left = Math.max(margin, Math.min(wantedSide, maxLeft));
      const maxTop = window.innerHeight - (menuRect ? menuRect.height : 0) - margin;
      const top = Math.max(margin, Math.min(anchorRect.top, maxTop));
      menuStyle.value = {
        position: "fixed",
        left: `${left}px`,
        top: `${top}px`,
      };
      return;
    }
    const wanted =
      align === "center"
        ? anchorRect.left + anchorRect.width / 2 - menuWidth / 2
        : anchorRect.left;
    const left = Math.max(margin, Math.min(wanted, maxLeft));
    // "above" flips below when the popup would run off the top of the
    // viewport — a status-bar control mounted at the top of a view
    // (SPEC-098's deck language control in the app bar) still opens on-screen.
    const above = anchorRect.top - offsetY - (menuRect ? menuRect.height : 0);
    const top = placement === "above" && above >= margin ? above : anchorRect.bottom + offsetY;
    menuStyle.value = {
      position: "fixed",
      left: `${left}px`,
      top: `${top}px`,
    };
  }

  // flush: "post" so the popup's DOM (gated on `open`) exists before its
  // width is measured.
  watch(open, (isOpen) => {
    if (isOpen) {
      position();
    }
  }, { flush: "post" });

  // Capture-phase scroll also picks up the toolbar's own horizontal scroll,
  // keeping the popup pinned to the button as the strip moves beneath it.
  document.addEventListener("scroll", position, true);
  window.addEventListener("resize", position);
  const instance = getCurrentInstance();
  if (instance) {
    onBeforeUnmount(() => {
      document.removeEventListener("scroll", position, true);
      window.removeEventListener("resize", position);
    });
  }

  return { menuStyle };
}
