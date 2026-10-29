<script setup lang="ts">
// A sheet drawing's context menu (SPEC-131): opened by a right-click on a
// drawing, or the Menu key / Shift+F10 on a focused one (useSheetDrawings'
// openContextMenu), at the pointer — `sheet-drawing-menu` with Bring to Front,
// Bring Forward, Send Backward and Send to Back (`sheet-drawing-front`,
// `-forward`, `-backward`, `-back`; the two that cannot move the drawing are
// disabled), then "Alt Text…" (`sheet-drawing-alt-text`) and Delete
// (`sheet-drawing-delete`). Each emits run(action) — "front" | "forward" |
// "backward" | "back" | "alt-text" | "delete" — after close. UiMenuSurface
// owns outside-click and Escape dismissal. The Arrange menu offers the same
// actions under its own `sheet-menu-arrange-*` ids (the menu bar's popups stay
// mounted, so the ids could not be shared). Runtime prop declarations per
// the espree constraint (see AppShell.vue).
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  open: { type: Boolean, default: false },
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
  /** The drawing is already at the front / the back of the z-order. */
  atFront: { type: Boolean, default: false },
  atBack: { type: Boolean, default: false },
  /** SPEC-138: no edit applies — every item is disabled. */
  readOnly: { type: Boolean, default: false },
});

const emit = defineEmits(["run", "close"]);

const ORDER = [
  { action: "front", label: "Bring to Front", testid: "sheet-drawing-front", end: "front" },
  { action: "forward", label: "Bring Forward", testid: "sheet-drawing-forward", end: "front" },
  { action: "backward", label: "Send Backward", testid: "sheet-drawing-backward", end: "back" },
  { action: "back", label: "Send to Back", testid: "sheet-drawing-back", end: "back" },
];

function choose(action = "front") {
  emit("close");
  emit("run", action);
}
</script>

<template>
  <UiMenuSurface
    :open="props.open"
    :anchor="null"
    class="sheet-drawing-menu"
    role="menu"
    aria-label="Drawing"
    data-testid="sheet-drawing-menu"
    :style="{ left: `${props.x}px`, top: `${props.y}px` }"
    @dismiss="emit('close')"
  >
    <button
      v-for="item in ORDER"
      :key="item.action"
      type="button"
      role="menuitem"
      :data-testid="item.testid"
      :disabled="props.readOnly || (item.end === 'front' ? props.atFront : props.atBack)"
      @mousedown.prevent
      @click="choose(item.action)"
    >
      {{ item.label }}
    </button>
    <hr
      class="menu-separator"
      role="separator"
    >
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-drawing-alt-text"
      :disabled="props.readOnly"
      @mousedown.prevent
      @click="choose('alt-text')"
    >
      Alt Text…
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-drawing-delete"
      :disabled="props.readOnly"
      @mousedown.prevent
      @click="choose('delete')"
    >
      Delete
    </button>
  </UiMenuSurface>
</template>

<style scoped>
.sheet-drawing-menu {
  position: fixed;
  z-index: 40;
  min-width: 11rem;
}

button {
  display: block;
  width: 100%;
  padding: var(--space-1) var(--space-3);
  font: inherit;
  text-align: left;
  color: var(--text-primary);
  background: transparent;
  border: 0;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

button:hover:not(:disabled),
button:focus-visible {
  background: var(--state-hover);
}

button:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.menu-separator {
  margin: var(--space-1) 0;
  border: 0;
  border-top: var(--border-width) solid var(--border-subtle);
}
</style>
