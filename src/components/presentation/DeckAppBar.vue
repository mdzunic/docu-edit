<script setup lang="ts">
// The deck's app-bar cluster: the deck facts (slide count · format) and the
// action row (undo/redo, Notes, Deck setup…, Present, Download, save state,
// Close). AppShell's #app-bar-actions slot is where the word processor already
// puts its own EditorSaveStatus, so slotting this here is what makes the two
// editors share one top line.
//
// It exists as a component rather than as markup in PresentationView because
// the view is a wiring hub under a line budget that only ratchets down. Like
// PresentationDialogs, it takes the controllers whole and calls them directly
// instead of re-emitting nine events the view would have to re-wire.
//
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import DeckActions from "./DeckActions.vue";

const props = defineProps({
  /** e.g. "4 slides" (SPEC-079's readout — presentations.spec asserts it). */
  slideCountLabel: { type: String, default: "" },
  /** e.g. "PowerPoint (.pptx)". */
  formatName: { type: String, default: "" },
  saveState: { type: String, default: "saved" },
  saveStateText: { type: String, default: "" },
  /** The useDeckActions controller (present / download / export / close). */
  deckActions: { type: Object, required: true },
  /** SPEC-104: the useDeckExport controller (the print-and-export dialog). */
  deckExport: { type: Object, default: null },
});

const c = props;
</script>

<template>
  <span
    class="deck-facts"
    data-testid="presentation-deck-facts"
  >
    <span data-testid="presentation-slide-count">{{ slideCountLabel }}</span>
    · {{ formatName }}
  </span>
  <DeckActions
    :save-state="saveState"
    :save-state-text="saveStateText"
    :can-present="c.deckActions.canPresent.value"
    :other-format-label="c.deckActions.otherFormatLabel.value"
    @present="c.deckActions.present"
    @download="c.deckActions.download()"
    @download-copy="c.deckActions.downloadCopy()"
    @export-html="c.deckActions.openHtmlExport()"
    @export-pdf="c.deckExport?.open('pdf')"
    @export-images="c.deckExport?.open('images')"
    @close="c.deckActions.close()"
  />
</template>

<style scoped>
/* The deck's facts beside the title, in the caption face the status bar uses —
   chrome to glance at, not a heading. It is the one thing on this bar that may
   shrink: a crowded bar should ellipsise "· POWERPOINT (.PPTX)" long before it
   pushes a command past the right edge. */
.deck-facts {
  flex: 0 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  white-space: nowrap;
}
</style>
