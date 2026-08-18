<script setup lang="ts">
import { shortcutLabel } from "../lib/shortcuts";

// Runtime prop/emit declarations, not the defineProps<{...}>()/defineEmits<{...}>()
// generics — see AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  canUndo: { type: Boolean, default: false },
  /** The action name the next undo would report, e.g. "typing" (SPEC-015
   * AC-7) — blank while canUndo is false. */
  undoLabel: { type: String, default: "" },
  canRedo: { type: Boolean, default: false },
  /** The action name the next redo would report (AC-8) — blank while
   * canRedo is false. */
  redoLabel: { type: String, default: "" },
});

const emit = defineEmits(["undo", "redo"]);

// Shortcut text is read from the SPEC-021 registry rather than hardcoded
// (AC-15, AC-16).
const undoShortcut = shortcutLabel("edit.undo");
const redoShortcut = shortcutLabel("edit.redo");

function undoTitle() {
  return props.canUndo ? `Undo ${props.undoLabel} (${undoShortcut})` : `Undo (${undoShortcut})`;
}

function redoTitle() {
  return props.canRedo ? `Redo ${props.redoLabel} (${redoShortcut})` : `Redo (${redoShortcut})`;
}
</script>

<template>
  <div
    class="undo-redo"
    role="group"
    aria-label="Undo and redo"
  >
    <button
      type="button"
      class="undo-redo-button"
      data-testid="undo-button"
      :disabled="!canUndo"
      :title="undoTitle()"
      :aria-label="undoTitle()"
      @mousedown.prevent
      @click="emit('undo')"
    >
      ↶
    </button>
    <button
      type="button"
      class="undo-redo-button"
      data-testid="redo-button"
      :disabled="!canRedo"
      :title="redoTitle()"
      :aria-label="redoTitle()"
      @mousedown.prevent
      @click="emit('redo')"
    >
      ↷
    </button>
  </div>
</template>

<style scoped>
.undo-redo {
  display: inline-flex;
  gap: var(--space-2);
}

.undo-redo-button {
  font: inherit;
  font-size: var(--text-lg);
  width: 1.75rem;
  height: 1.75rem;
  padding: 0;
  line-height: 1;
  text-align: center;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.undo-redo-button:disabled {
  opacity: 0.5;
  cursor: default;
}
</style>
