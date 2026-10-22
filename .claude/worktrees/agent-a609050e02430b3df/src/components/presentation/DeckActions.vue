<script setup lang="ts">
// The deck's app-bar action cluster (Present, download, save state, Close) —
// extracted from PresentationView so the view stays a thin wiring hub (its size
// budget only ratchets down). Pure presentation + event forwarding; the view
// owns the commands these fire.
//
// It is deliberately the word processor's cluster plus Present: save state,
// the download control, Close. Four controls left when the deck line folded
// into the app bar, because eight of them overflowed the bar at 1024px and put
// Close off the right edge where nothing could reach it — and each had a truer
// home anyway:
//
//   · Notes was never an action. It opens the notes panel exactly as Design,
//     Table and Media open theirs, so it joined SlideRibbon's panel toggles
//     (its slides-notes-toggle testid unchanged).
//   · Undo/redo went to the ribbon row, which is the deck's always-visible
//     command line — where DocumentToolbar's band 1 keeps the word processor's
//     pair, and where the handoff's own 2c strip starts.
//   · Deck setup… is a rare command the File menu already carries
//     (slides-menu-deck-setup); the handoff's 2c bar holds only Download,
//     Present and Close.
//
// Close is now the same icon button the word processor's cluster ends with,
// rather than a wider text button saying the same thing.
//
// Runtime prop/emit declarations, not the generic forms (espree lints .vue
// scripts).
import SaveStateBadge from "../SaveStateBadge.vue";
import UiButton from "../ui/UiButton.vue";
import UiIconButton from "../ui/UiIconButton.vue";
import PresentationDownloadButton from "./PresentationDownloadButton.vue";

defineProps({
  /** SPEC-010's readout, for the deck. The word processor has shown this all
   * along and the slides side showed nothing, which left the product's own
   * "stored in this browser" promise uncheckable on half of it. */
  saveState: { type: String, default: "saved" },
  saveStateText: { type: String, default: "" },
  canPresent: { type: Boolean, default: false },
  otherFormatLabel: { type: String, default: "" },
});

const emit = defineEmits([
  "present",
  "download",
  "download-copy",
  "export-html",
  "export-pdf",
  "export-images",
  "close",
]);
</script>

<template>
  <div class="deck-actions">
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
      @export-pdf="emit('export-pdf')"
      @export-images="emit('export-images')"
    />
    <SaveStateBadge
      v-if="saveStateText"
      class="deck-save-state"
      :save-state="saveState"
      :text="saveStateText"
    />
    <UiIconButton
      class="close-button"
      icon="close"
      label="Close presentation"
      data-testid="presentation-close"
      @click="emit('close')"
    />
  </div>
</template>

<style scoped>
.deck-actions {
  display: flex;
  flex: none;
  gap: var(--space-3);
  align-items: center;
}

/* Quieter than the buttons around it: a state to glance at, not an action. */
.deck-save-state {
  color: var(--text-secondary);
}

/* The icon button's default greys are tuned for light surfaces; on the brand
   bar it takes the on-brand foreground and a translucent wash so it reads in
   both themes — the same treatment EditorSaveStatus gives the word processor's
   Close, because it is now the same control. */
.deck-actions .close-button {
  color: var(--text-on-brand);
}

.deck-actions .close-button:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.deck-actions .close-button:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}
</style>
