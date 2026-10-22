<script setup lang="ts">
// The image editor's action bar (SPEC-112): Undo/Redo in SPEC-015's
// vocabulary (the tooltip names the step: "Undo Crop"), Revert to original,
// Flatten and discard original, Export, Save as a copy, Save and Done.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { asImageEditor } from "../../views/media/useImageEditor";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const ready = computed(() => e.value.phase.value === "ready" && !e.value.actions.busy.value);
const undoLabel = computed(() => (e.value.undo.undoLabel.value ? `Undo ${e.value.undo.undoLabel.value}` : "Undo"));
const redoLabel = computed(() => (e.value.undo.redoLabel.value ? `Redo ${e.value.undo.redoLabel.value}` : "Redo"));
</script>

<template>
  <div
    class="edit-bar"
    role="toolbar"
    aria-label="Image editor actions"
  >
    <UiButton
      size="sm"
      variant="ghost"
      data-testid="edit-undo"
      :title="undoLabel"
      :aria-label="undoLabel"
      :disabled="!ready || !e.undo.canUndo.value"
      @click="e.undoStep()"
    >
      Undo
    </UiButton>
    <UiButton
      size="sm"
      variant="ghost"
      data-testid="edit-redo"
      :title="redoLabel"
      :aria-label="redoLabel"
      :disabled="!ready || !e.undo.canRedo.value"
      @click="e.redoStep()"
    >
      Redo
    </UiButton>
    <span
      class="divider"
      aria-hidden="true"
    />
    <UiButton
      size="sm"
      data-testid="edit-revert"
      :disabled="!ready || e.history.value.length === 0"
      @click="e.requestRevert()"
    >
      Revert to original
    </UiButton>
    <UiButton
      size="sm"
      data-testid="edit-flatten"
      :disabled="!ready || e.stack.value.length === 0"
      @click="e.actions.requestFlatten()"
    >
      Flatten and discard original
    </UiButton>
    <span class="spacer" />
    <UiButton
      size="sm"
      data-testid="edit-export"
      :disabled="!ready"
      @click="e.actions.openExport()"
    >
      Export…
    </UiButton>
    <UiButton
      size="sm"
      data-testid="edit-save-copy"
      :disabled="!ready"
      @click="e.actions.saveCopy()"
    >
      Save as a copy
    </UiButton>
    <UiButton
      size="sm"
      variant="primary"
      data-testid="edit-save"
      :disabled="!ready"
      @click="e.actions.save()"
    >
      Save
    </UiButton>
    <UiButton
      size="sm"
      data-testid="edit-done"
      :disabled="e.actions.busy.value"
      @click="e.done()"
    >
      Done
    </UiButton>
  </div>
</template>

<style scoped>
.edit-bar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
}

.divider {
  width: var(--border-width);
  align-self: stretch;
  background: var(--border-subtle);
}

.spacer {
  flex: 1 1 auto;
}
</style>
