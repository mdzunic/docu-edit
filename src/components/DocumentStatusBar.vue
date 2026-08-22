<script setup lang="ts">
// The editor's status bar (SPEC-063 extraction). Bundles the page
// indicator, the transient opening/paginating notices, the word-count
// button, the zoom control (AC-1: after the page indicator, before the word
// count), the language control and the track-changes toggle into one
// component so DocumentView's template holds a single wiring tag rather than
// the whole strip. The controls themselves are unchanged; only their host
// moved.
import LanguageStatusControl from "./LanguageStatusControl.vue";
import ZoomControl from "./ZoomControl.vue";

// Runtime prop/emit declarations (espree bars TS-only syntax in .vue
// scripts — see CLAUDE.md).
defineProps({
  caretPage: { type: Number, default: 0 },
  pageCount: { type: Number, default: 0 },
  documentOpening: { type: Boolean, default: false },
  paginationBusy: { type: Boolean, default: false },
  wordCountText: { type: String, default: "" },
  zoomPercent: { type: Number, default: 100 },
  zoomFitMode: { type: String, default: "none" },
  currentLanguageTag: { type: String, default: "" },
  currentLanguageLabel: { type: String, default: "" },
  languageOptions: { type: Array, default: () => [] },
  languageSelectionActive: { type: Boolean, default: false },
  selectionNoProof: { type: Boolean, default: false },
  editable: { type: Boolean, default: false },
  trackChangesRecording: { type: Boolean, default: false },
});

const emit = defineEmits([
  "open-statistics",
  "manage-dictionary",
  "toggle-track-changes",
  "zoom-in",
  "zoom-out",
  "zoom-fit-width",
  "zoom-fit-page",
  "set-language",
  "toggle-no-proof",
  "zoom-apply-input",
  "zoom-set-preset",
]);
</script>

<template>
  <span data-testid="page-indicator">Page {{ caretPage }} of {{ pageCount }}</span>
  <span
    v-if="documentOpening"
    role="status"
    data-testid="document-opening"
  >Opening…</span>
  <span
    v-if="paginationBusy"
    data-testid="pagination-busy"
  >Paginating…</span>
  <ZoomControl
    :zoom-percent="zoomPercent"
    :fit-mode="zoomFitMode"
    @zoom-in="emit('zoom-in')"
    @zoom-out="emit('zoom-out')"
    @apply-input="emit('zoom-apply-input', $event)"
    @set-preset="emit('zoom-set-preset', $event)"
    @fit-width="emit('zoom-fit-width')"
    @fit-page="emit('zoom-fit-page')"
  />
  <button
    type="button"
    class="word-count-button"
    data-testid="word-count-button"
    @click="emit('open-statistics')"
  >
    {{ wordCountText }}
  </button>
  <LanguageStatusControl
    :current-language-tag="currentLanguageTag"
    :current-language-label="currentLanguageLabel"
    :language-options="languageOptions"
    :selection-active="languageSelectionActive"
    :selection-no-proof="selectionNoProof"
    @select="emit('set-language', $event)"
    @manage-dictionary="emit('manage-dictionary')"
    @toggle-no-proof="emit('toggle-no-proof', $event)"
  />
  <button
    v-if="editable"
    type="button"
    class="track-changes-status-indicator"
    data-testid="status-bar-track-changes"
    :aria-pressed="trackChangesRecording"
    @click="emit('toggle-track-changes')"
  >
    Track changes: {{ trackChangesRecording ? "On" : "Off" }}
  </button>
</template>

<style scoped>
/* A plain-looking button (no border/background of its own) so it reads as
   status text until focused/hovered, while still being a real <button>
   (SPEC-036 AC-1, AC-4, AC-27). */
.word-count-button,
.track-changes-status-indicator {
  font: inherit;
  font-size: var(--text-sm);
  color: inherit;
  padding: 0;
  border: none;
  background: none;
  cursor: pointer;
}

.word-count-button:hover,
.word-count-button:focus-visible {
  text-decoration: underline;
}

/* SPEC-039 AC-1 / SPEC-061 AC-19: the status-bar control is the single
   track-changes toggle, matching word-count-button's plain-until-hovered
   look. */
.track-changes-status-indicator[aria-pressed="true"] {
  text-decoration: underline;
  font-weight: var(--weight-bold);
}
</style>
