<script setup lang="ts">
// SPEC-125: the fill options chip (`sheet-fill-options`, role="toolbar"),
// shown at a drag fill's bottom-right corner: Copy cells, Fill series, Fill
// formatting only and Fill without formatting
// (`sheet-fill-copy|series|formats|no-formats`), the applied one pressed.
// Choosing one emits choose(mode); useFillHandle re-applies the fill as the
// same undo step, and dismisses the chip on Escape or any other command. A
// press here never reaches the grid (no selection, focus kept). Runtime prop
// declarations per the espree constraint (see AppShell.vue).
const props = defineProps({
  /** Paper px of the filled range's bottom-right corner. */
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
  /** "copy" | "series" | "formats" | "no-formats". */
  applied: { type: String, default: "series" },
});

const emit = defineEmits(["choose"]);

const OPTIONS = [
  { mode: "copy", label: "Copy cells", testid: "sheet-fill-copy" },
  { mode: "series", label: "Fill series", testid: "sheet-fill-series" },
  { mode: "formats", label: "Fill formatting only", testid: "sheet-fill-formats" },
  { mode: "no-formats", label: "Fill without formatting", testid: "sheet-fill-no-formats" },
];
</script>

<template>
  <div
    class="sheet-fill-options"
    role="toolbar"
    aria-label="Fill options"
    data-testid="sheet-fill-options"
    :style="{ left: `${props.x}px`, top: `${props.y}px` }"
    @pointerdown.stop
    @mousedown.prevent
    @dblclick.stop
  >
    <button
      v-for="option in OPTIONS"
      :key="option.mode"
      type="button"
      :data-testid="option.testid"
      :aria-pressed="props.applied === option.mode ? 'true' : 'false'"
      @click="emit('choose', option.mode)"
    >
      {{ option.label }}
    </button>
  </div>
</template>

<style scoped>
.sheet-fill-options {
  position: absolute;
  z-index: 5;
  display: flex;
  gap: var(--space-1);
  margin: var(--space-1) 0 0 var(--space-1);
  padding: var(--space-1);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-md, 0 2px 6px rgb(0 0 0 / 20%));
  white-space: nowrap;
}

button {
  padding: var(--space-1) var(--space-2);
  font: inherit;
  font-size: var(--text-xs);
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

button:hover,
button:focus-visible {
  background: var(--state-hover);
}

button[aria-pressed="true"] {
  border-color: var(--accent);
  background: color-mix(in srgb, var(--accent) 14%, transparent);
}
</style>
