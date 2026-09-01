<script setup lang="ts">
// The deck workspace body (SPEC-093 extraction, the DeckHeader/SlidePanelStack
// precedent): the filmstrip, the fit-scaled stage and the below-stage panel
// stack — the block below the editing bars. Extracted from PresentationView.vue
// to pay for the Header & Footer wiring while keeping the view a thin hub
// (PresentationView.size.spec.ts). A dumb component: every controller/command it
// needs is passed in (as reactive props read off the deck's shallowRef, so an
// in-place mutation + triggerRef still updates the filmstrip/stage), and it
// re-emits the filmstrip intents unchanged. Runtime prop declarations, not the
// generic forms (espree lints .vue scripts).
import { computed } from "vue";
import SlideChartLayer from "./SlideChartLayer.vue";
import SlideFilmstrip from "./SlideFilmstrip.vue";
import SlideStageColumn from "./SlideStageColumn.vue";
import { asStageFieldsBySlide } from "./slideRender";

const props = defineProps({
  slides: { type: Array, default: () => [] },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  currentSlide: { type: Object, default: null },
  currentIndex: { type: Number, default: 0 },
  media: { type: Array, default: () => [] },
  /** SPEC-094: the deck's theme, threaded to the stage surfaces. */
  theme: { type: Object, default: null },
  liveLabel: { type: String, default: "" },
  stageEdit: { type: Object, required: true },
  animations: { type: Object, required: true },
  transitions: { type: Object, required: true },
  notes: { type: Object, required: true },
  tables: { type: Object, required: true },
  charts: { type: Object, required: true },
  /** SPEC-093: each slide's pre-resolved deck field band, indexed to `slides`. */
  fieldsBySlide: { type: Array, default: () => [] },
});

defineEmits(["select", "add", "command", "reorder"]);

const currentFields = computed(() => asStageFieldsBySlide(props.fieldsBySlide)[props.currentIndex] ?? []);
</script>

<template>
  <SlideChartLayer :controller="charts" />
  <div class="deck-body">
    <SlideFilmstrip
      :slides="slides"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :current-index="currentIndex"
      :fields-by-slide="fieldsBySlide"
      :theme="theme"
      @select="$emit('select', $event)"
      @add="$emit('add')"
      @command="$emit('command', $event)"
      @reorder="$emit('reorder', $event)"
    />
    <!-- The .stage-column block lives in SlideStageColumn (SPEC-096) -->
    <SlideStageColumn
      :current-slide="currentSlide ?? undefined"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :live-label="liveLabel"
      :stage-edit="stageEdit"
      :theme="theme"
      :animations="animations"
      :transitions="transitions"
      :notes="notes"
      :tables="tables"
      :current-index="currentIndex"
      :fields="currentFields"
      @add="$emit('add')"
    />
  </div>
</template>

<style scoped>
.deck-body {
  flex: 1 1 auto;
  min-height: 0;
  display: grid;
  grid-template-columns: minmax(180px, 220px) 1fr;
}
.deck-body :deep(.filmstrip) {
  border-right: var(--border-width) solid var(--border-subtle);
  max-height: 100%;
}
</style>
