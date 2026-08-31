<script setup lang="ts">
// The deck workspace body (SPEC-093 extraction, the DeckHeader/SlidePanelStack
// precedent): the filmstrip, the fit-scaled stage and the below-stage panel
// stack — the block below the editing bars. Extracted from PresentationView.vue
// to pay for the Header & Footer wiring while keeping the view a thin hub
// (PresentationView.size.spec.ts). A dumb component: every controller/command it
// needs is passed in, and it re-emits the filmstrip intents unchanged. Runtime
// prop declarations, not the generic forms (espree lints .vue scripts).
import { computed } from "vue";
import SlideChartLayer from "./SlideChartLayer.vue";
import SlideFilmstrip from "./SlideFilmstrip.vue";
import SlideStageArea from "./SlideStageArea.vue";
import SlidePanelStack from "./SlidePanelStack.vue";
import { asPresentationModel } from "./slideRender";

const props = defineProps({
  model: { type: Object, default: null },
  currentIndex: { type: Number, default: 0 },
  media: { type: Array, default: () => [] },
  liveLabel: { type: String, default: "" },
  stageEdit: { type: Object, required: true },
  animations: { type: Object, required: true },
  transitions: { type: Object, required: true },
  notes: { type: Object, required: true },
  tables: { type: Object, required: true },
  charts: { type: Object, required: true },
});

defineEmits(["select", "add", "command", "reorder"]);

const slides = computed(() => asPresentationModel(props.model)?.slides ?? []);
const widthCm = computed(() => asPresentationModel(props.model)?.slideWidthCm ?? 0);
const heightCm = computed(() => asPresentationModel(props.model)?.slideHeightCm ?? 0);
const currentSlide = computed(() => slides.value[props.currentIndex] ?? null);
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
      :model="model"
      @select="$emit('select', $event)"
      @add="$emit('add')"
      @command="$emit('command', $event)"
      @reorder="$emit('reorder', $event)"
    />
    <div class="stage-column">
      <SlideStageArea
        :current-slide="currentSlide ?? undefined"
        :width-cm="widthCm"
        :height-cm="heightCm"
        :media="media"
        :live-label="liveLabel"
        :stage-edit="stageEdit"
        :animations="animations"
        :transitions="transitions"
        :model="model"
        :slide-index="currentIndex"
        @add="$emit('add')"
      />
      <SlidePanelStack
        :notes="notes"
        :transitions="transitions"
        :animations="animations"
        :tables="tables"
        :slide-number="currentIndex + 1"
        :has-slide="!!currentSlide"
      />
    </div>
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
.stage-column {
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
}
</style>
