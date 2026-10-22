<script setup lang="ts">
import { onBeforeUnmount, useTemplateRef, watch } from "vue";

// Popup container primitive (SPEC-058). This is the one place popup dismissal
// is implemented, so no dropdown re-rolls it: elevation, radius and padding
// from tokens, plus outside-click AND Escape dismissal wired to the *document*
// (not the surface root). That distinction is the whole point — when a popup
// is opened by a mouse click, focus stays on the document, so a root-level
// @keydown.escape never fires and Escape silently does nothing. Listening on
// document fixes keyboard dismissal for mouse-opened menus. On close, focus
// returns to the anchor that opened the surface.
//
// ARIA roles stay the consumer's choice (menu / listbox / grid) — this only
// owns the box and the dismissal, via the default slot.
const props = defineProps({
  /** Whether the surface is shown. */
  open: { type: Boolean, default: false },
  /** The trigger element: excluded from outside-click, and focus returns to
   * it on close. May be an HTMLElement or a component/ref exposing $el. */
  anchor: { type: null, default: null },
  /** Set false when the consumer runs its own focus discipline (MenuBar's
   * roving tab stop): the surface then never moves focus on close, so it
   * cannot fight the consumer over where focus lands on a pointer
   * cross-open or an outside click. */
  returnFocus: { type: Boolean, default: true },
  /** Keep the closed surface in the DOM (hidden) instead of unmounting it.
   * MenuBar needs this: its popups' contents are part of the page's stated
   * contract while closed — SPEC-020 reads the format painter's
   * aria-checked off a closed Format menu — and slotted widgets keep their
   * mounted state across open/close. Everything else (listeners, focus
   * return) still keys off `open`. */
  keepMounted: { type: Boolean, default: false },
});

const emit = defineEmits(["dismiss"]);

// The consumer positions the surface (useFloatingMenu inline style) and owns
// its ARIA role + data-testid; those arrive as fall-through attrs and are
// bound explicitly onto the box so nothing leaks when the surface is closed
// (no "extraneous attributes" warning on the v-if comment node).
defineOptions({ inheritAttrs: false });

const surface = useTemplateRef("surface");

function anchorEl() {
  const a = props.anchor;
  if (a instanceof HTMLElement) {
    return a;
  }
  if (a && typeof a === "object" && a.$el instanceof HTMLElement) {
    return a.$el;
  }
  return null;
}

function onDocPointerDown(event = new PointerEvent("pointerdown")) {
  const target = event.target;
  if (!(target instanceof Node)) {
    return;
  }
  if (surface.value instanceof HTMLElement && surface.value.contains(target)) {
    return;
  }
  // A pointerdown inside ANY menu surface leaves this one open. A surface can
  // spawn a child surface that is a DOM sibling (StylePicker's per-entry
  // context menu), and dismissing the parent here unmounts the child between
  // pointerdown and click — the item's click never fires, so the action
  // silently does nothing (SPEC-019's update/reset-style regressions). Two
  // unrelated surfaces can never both be open: opening the second one starts
  // with a pointerdown outside the first, which dismisses it right here.
  if (target instanceof Element && target.closest(".ui-menu-surface")) {
    return;
  }
  const anchor = anchorEl();
  if (anchor && anchor.contains(target)) {
    // Let the trigger's own click toggle the menu closed; dismissing here too
    // would immediately re-open it.
    return;
  }
  emit("dismiss");
}

function onDocKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    event.stopPropagation();
    emit("dismiss");
  }
}

function addListeners() {
  document.addEventListener("pointerdown", onDocPointerDown, true);
  document.addEventListener("keydown", onDocKeydown);
}

function removeListeners() {
  document.removeEventListener("pointerdown", onDocPointerDown, true);
  document.removeEventListener("keydown", onDocKeydown);
}

watch(
  () => props.open,
  (isOpen, wasOpen) => {
    if (isOpen) {
      addListeners();
      return;
    }
    removeListeners();
    // Return focus to the opener when the surface closes, but only if focus is
    // still inside the surface (or nowhere) — never steal it from a control the
    // user has already deliberately moved to.
    if (wasOpen && props.returnFocus) {
      const active = document.activeElement;
      const inside =
        surface.value instanceof HTMLElement &&
        active instanceof Node &&
        surface.value.contains(active);
      if (inside || active === document.body || active === null) {
        anchorEl()?.focus();
      }
    }
  },
);

onBeforeUnmount(removeListeners);
</script>

<template>
  <div
    v-if="props.open || props.keepMounted"
    v-show="props.open"
    ref="surface"
    class="ui-menu-surface"
    v-bind="$attrs"
  >
    <slot />
  </div>
</template>

<style scoped>
.ui-menu-surface {
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  padding: var(--space-2);
}
</style>
