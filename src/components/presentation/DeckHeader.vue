<script setup lang="ts">
// The presentation deck header — the deck title/facts line plus the
// DeckActions cluster — extracted from PresentationView (SPEC-090) so the view
// stays a thin wiring hub within its size budget. Pure presentation + event
// forwarding; the view owns the commands these fire. Runtime prop/emit
// declarations, not the generic forms (espree lints .vue scripts).
import DeckActions from "./DeckActions.vue";

defineProps({
  saveState: { type: String, default: "saved" },
  saveStateText: { type: String, default: "" },
  deckTitle: { type: String, default: "" },
  slideCountLabel: { type: String, default: "" },
  formatName: { type: String, default: "" },
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
  <header class="deck-header">
    <div class="deck-facts">
      <h2 class="deck-title">
        {{ deckTitle }}
      </h2>
      <p class="deck-meta">
        <span data-testid="presentation-slide-count">{{ slideCountLabel }}</span>
        · {{ formatName }}
      </p>
    </div>
    <DeckActions
      :save-state="saveState"
      :save-state-text="saveStateText"
      :can-undo="canUndo"
      :can-redo="canRedo"
      :notes-open="notesOpen"
      :can-present="canPresent"
      :other-format-label="otherFormatLabel"
      @undo="emit('undo')"
      @redo="emit('redo')"
      @toggle-notes="emit('toggle-notes')"
      @deck-setup="emit('deck-setup')"
      @present="emit('present')"
      @download="emit('download')"
      @download-copy="emit('download-copy')"
      @export-html="emit('export-html')"
      @close="emit('close')"
    />
  </header>
</template>

<style scoped>
.deck-header {
  display: flex;
  align-items: flex-start;
  gap: var(--space-4);
  flex-wrap: wrap;
  padding: var(--space-4) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.deck-facts {
  flex: 1 1 auto;
  min-width: 0;
}

.deck-title {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.deck-meta {
  margin: var(--space-1) 0 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}
</style>
