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
    if (wasOpen) {
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
    v-if="props.open"
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
