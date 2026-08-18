<script setup lang="ts">
// The paste options chip (SPEC-016 AC-29..33): floats near the caret after
// a paste and re-interprets that paste as "Keep source formatting", "Match
// destination formatting" or "Keep text only". Emits the chosen mode only —
// the editor owns the re-application (and its single-undo-step contract).
// Runtime prop declarations, not the defineProps<{...}>() generics — see
// AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  /** Viewport coordinates the chip is pinned at (below the pasted caret). */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  /** The interpretation currently applied, shown pressed. */
  applied: { type: String, default: "source" },
});

const emit = defineEmits(["select", "dismiss"]);

const OPTIONS = [
  { mode: "source", label: "Keep source formatting" },
  { mode: "match", label: "Match destination formatting" },
  { mode: "text", label: "Keep text only" },
];

function chipStyle() {
  return { left: `${props.left}px`, top: `${props.top}px` };
}
</script>

<template>
  <div
    class="paste-chip"
    data-testid="paste-options-chip"
    role="toolbar"
    aria-label="Paste options"
    :style="chipStyle()"
    @keydown.escape.stop="emit('dismiss')"
  >
    <button
      v-for="option in OPTIONS"
      :key="option.mode"
      type="button"
      class="paste-chip-option"
      :class="{ applied: option.mode === applied }"
      :data-testid="`paste-option-${option.mode}`"
      :aria-pressed="option.mode === applied ? 'true' : 'false'"
      @mousedown.prevent
      @click="emit('select', option.mode)"
    >
      {{ option.label }}
    </button>
  </div>
</template>

<style scoped>
.paste-chip {
  position: fixed;
  z-index: 30;
  display: flex;
  gap: var(--space-2);
  padding: var(--space-2);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.paste-chip-option {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  white-space: nowrap;
}

.paste-chip-option:hover {
  background: var(--surface-canvas);
}

.paste-chip-option.applied {
  background: var(--surface-selected);
  border-color: var(--accent-subtle);
  color: var(--text-accent);
}
</style>
