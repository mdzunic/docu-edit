<script setup lang="ts">
// The row and column header menu (SPEC-124 "Header menu"): opened by a
// right-click on a header (`sheet-row-header`/`sheet-col-header`), or by
// Shift+F10 / the Menu key when the selection is whole rows or columns
// (useStructure), at the pointer or under the active band's header. Its
// entries — insert before/after, delete, hide, unhide, size, autofit — and
// their labels come from headerMenu.ts; choosing one emits its command.
// UiMenuSurface owns outside-click and Escape dismissal. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  open: { type: Boolean, default: false },
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
  /** `[{ command, label, testid, disabled }]` (headerMenu.ts). */
  items: { type: Array, default: () => [] },
});

const emit = defineEmits(["run", "close"]);

function entries() {
  return props.items.map((item) => Object(item));
}

function choose(command = "") {
  emit("run", command);
}
</script>

<template>
  <UiMenuSurface
    :open="props.open"
    :anchor="null"
    class="sheet-header-menu"
    role="menu"
    aria-label="Rows and columns"
    data-testid="sheet-header-menu"
    :style="{ left: `${props.x}px`, top: `${props.y}px` }"
    @dismiss="emit('close')"
  >
    <button
      v-for="item in entries()"
      :key="item.command"
      type="button"
      role="menuitem"
      :data-testid="item.testid"
      :disabled="item.disabled"
      :aria-disabled="item.disabled ? 'true' : undefined"
      @mousedown.prevent
      @click="choose(item.command)"
    >
      {{ item.label }}
    </button>
  </UiMenuSurface>
</template>

<style scoped>
.sheet-header-menu {
  position: fixed;
  z-index: 40;
  min-width: 12rem;
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
  color: var(--text-secondary);
  cursor: default;
}
</style>
