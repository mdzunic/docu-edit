<script setup lang="ts">
import { modifierKeyLabel } from "../lib/platform";

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

const mod = modifierKeyLabel();

function undoTitle() {
  return props.canUndo ? `Undo ${props.undoLabel} (${mod}+Z)` : `Undo (${mod}+Z)`;
}

function redoTitle() {
  return props.canRedo ? `Redo ${props.redoLabel} (${mod}+Shift+Z)` : `Redo (${mod}+Shift+Z)`;
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
  gap: 0.25rem;
}

.undo-redo-button {
  font: inherit;
  font-size: 1rem;
  width: 1.75rem;
  height: 1.75rem;
  padding: 0;
  line-height: 1;
  text-align: center;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.undo-redo-button:disabled {
  opacity: 0.5;
  cursor: default;
}
</style>
