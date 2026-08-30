<script setup lang="ts">
// The deck-header action cluster (undo/redo, Notes toggle, Present, download,
// Close) — extracted from PresentationView so the view stays a thin wiring hub
// (its size budget only ratchets down). Pure presentation + event forwarding;
// the view owns the commands these fire. Runtime prop/emit declarations, not
// the generic forms (espree lints .vue scripts).
import AppIcon from "../AppIcon.vue";
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
  "deck-setup",
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
      title="Undo"
      @click="emit('undo')"
    >
      <AppIcon name="undo" />
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="presentation-redo"
      :disabled="!canRedo"
      aria-label="Redo"
      title="Redo"
      @click="emit('redo')"
    >
      <AppIcon name="redo" />
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
      variant="ghost"
      data-testid="presentation-deck-setup"
      @click="emit('deck-setup')"
    >
      Deck setup…
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
