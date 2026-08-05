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
  gap: 0.25rem;
  padding: 0.25rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.paste-chip-option {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.625rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  white-space: nowrap;
}

.paste-chip-option:hover {
  background: #f1f3f4;
}

.paste-chip-option.applied {
  background: #e8f0fe;
  border-color: #aecbfa;
  color: #174ea6;
}
</style>
