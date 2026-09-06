<script setup lang="ts">
import { shortcutLabel } from "../lib/shortcuts";
import UiIconButton from "./ui/UiIconButton.vue";

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
    <UiIconButton
      icon="undo"
      caption="Undo"
      :label="undoTitle()"
      :disabled="!canUndo"
      data-testid="undo-button"
      @mousedown.prevent
      @click="emit('undo')"
    />
    <UiIconButton
      icon="redo"
      caption="Redo"
      :label="redoTitle()"
      :disabled="!canRedo"
      data-testid="redo-button"
      @mousedown.prevent
      @click="emit('redo')"
    />
  </div>
</template>

<style scoped>
.undo-redo {
  display: inline-flex;
  gap: var(--space-2);
}
</style>
