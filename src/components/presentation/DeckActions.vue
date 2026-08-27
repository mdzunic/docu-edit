<script setup lang="ts">
// The deck-header action cluster (undo/redo, Notes toggle, Present, download,
// Close) — extracted from PresentationView so the view stays a thin wiring hub
// (its size budget only ratchets down). Pure presentation + event forwarding;
// the view owns the commands these fire. Runtime prop/emit declarations, not
// the generic forms (espree lints .vue scripts).
import UiButton from "../ui/UiButton.vue";
import PresentationDownloadButton from "./PresentationDownloadButton.vue";

defineProps({
  canUndo: { type: Boolean, default: false },
  canRedo: { type: Boolean, default: false },
  notesOpen: { type: Boolean, default: false },
  canPresent: { type: Boolean, default: false },
  otherFormatLabel: { type: String, default: "" },
});

const emit = defineEmits([
  "undo",
  "redo",
  "toggle-notes",
  "present",
  "download",
  "download-copy",
  "export-html",
  "close",
]);
</script>

<template>
  <div class="deck-actions">
    <UiButton
      variant="ghost"
      data-testid="presentation-undo"
      :disabled="!canUndo"
      aria-label="Undo"
      @click="emit('undo')"
    >
      Undo
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="presentation-redo"
      :disabled="!canRedo"
      aria-label="Redo"
      @click="emit('redo')"
    >
      Redo
    </UiButton>
    <UiButton
      variant="ghost"
      class="notes-toggle"
      data-testid="slides-notes-toggle"
      :aria-pressed="notesOpen"
      @click="emit('toggle-notes')"
    >
      Notes
    </UiButton>
    <UiButton
      variant="primary"
      data-testid="slides-present"
      :disabled="!canPresent"
      :title="canPresent ? 'Start slideshow' : 'Add a slide to present'"
      @click="emit('present')"
    >
      Present
    </UiButton>
    <PresentationDownloadButton
      :other-format-label="otherFormatLabel"
      @download="emit('download')"
      @download-copy="emit('download-copy')"
      @export-html="emit('export-html')"
    />
    <UiButton
      variant="ghost"
      data-testid="presentation-close"
      @click="emit('close')"
    >
      Close
    </UiButton>
  </div>
</template>

<style scoped>
.deck-actions {
  display: flex;
  gap: var(--space-3);
  align-items: center;
}
</style>
