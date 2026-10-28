<script setup lang="ts">
// The deck's version history panel (SPEC-103) — SPEC-040's pane for slides,
// mounted in the stage column under the stage like the other below-stage
// panels. Since SPEC-138 it wraps the product-neutral HistoryPanel.vue
// (`slides-history-panel`, every `history-*` testid unchanged) with the
// deck's own facts: the slide count and change, the .pptx/.odp download, the
// restore note — and its preview, the reconstructed deck drawn read-only on
// its own SlideStage with its own slide strip, so the live deck and its stage
// are never swapped (no diff highlighting).
//
// SPEC-085/086's panel contract: all logic lives in the useDeckHistory
// controller passed in, this component emits nothing, and its typed
// transforms live in deckHistoryRender.ts. Runtime prop declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import { computed } from "vue";
import HistoryPanel from "../HistoryPanel.vue";
import SlideStage from "./SlideStage.vue";
import {
  asDeckHistory,
  deckPanelController,
  previewFields,
  previewScale,
  previewSlide,
  previewTitle,
  RESTORE_UNDO_NOTE,
} from "./deckHistoryRender";

const props = defineProps({
  /** The useDeckHistory controller (useDeckSaving's `history`). */
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});

const c = asDeckHistory(props.controller);
const panel = deckPanelController(c);
const preview = computed(() => c.previewModel.value);
const shownSlide = computed(() => previewSlide(preview.value, c.previewIndex.value));
</script>

<template>
  <HistoryPanel
    class="deck-history"
    :controller="panel"
    testid="slides-history-panel"
    :label="`Version history (slide ${slideNumber})`"
    :restore-note="RESTORE_UNDO_NOTE"
    :preview-label="`Previewing ${previewTitle(c.versions.value, c.previewVersionId.value)} — read-only`"
    exit-label="Back to current deck"
  >
    <template #preview>
      <template v-if="preview">
        <SlideStage
          v-if="shownSlide"
          class="history-preview-stage"
          :slide="shownSlide"
          :width-cm="preview.slideWidthCm"
          :height-cm="preview.slideHeightCm"
          :media="preview.media ?? []"
          :scale="previewScale(preview)"
          :interactive="false"
          :fields="previewFields(preview, c.previewIndex.value)"
          :theme="preview.theme"
        />
        <div
          class="history-preview-strip"
          role="group"
          aria-label="Preview slides"
        >
          <button
            v-for="(slide, i) in preview.slides"
            :key="slide.id"
            type="button"
            class="history-preview-slide"
            :aria-pressed="i === c.previewIndex.value ? 'true' : 'false'"
            :aria-label="`Preview slide ${i + 1}`"
            @click="c.selectPreviewSlide(i)"
          >
            {{ i + 1 }}
          </button>
        </div>
      </template>
    </template>
  </HistoryPanel>
</template>

<style scoped>
/* Below the stage: a bounded strip of the column, ruled off from it. */
.deck-history {
  flex: 0 0 auto;
  border-top: var(--border-width) solid var(--border-subtle);
  max-height: 45%;
}

/* The preview is a picture of the version: nothing on it takes input. */
.history-preview-stage {
  pointer-events: none;
}

.history-preview-strip {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-1);
}

.history-preview-slide {
  min-width: 2rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.history-preview-slide[aria-pressed="true"] {
  border-color: var(--accent);
  background: var(--surface-selected);
}
</style>
