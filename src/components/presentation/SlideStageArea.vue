<script setup lang="ts">
// The scaling stage workspace (extracted from PresentationView for SPEC-089's
// line budget): the fit-scaled SlideStage, the empty-deck state, and the
// animation/transition preview layers. It owns the fit-to-workspace scaling (a
// ResizeObserver + the letterbox maths that were in the view), reacting to slide
// and deck-size changes; all editing wiring is threaded through the
// `stageEdit`/`animations`/`transitions` controllers so the parent stays a thin
// wiring hub. Espree lints .vue — runtime declarations, no TS generics.
import { onBeforeUnmount, onMounted, ref, useTemplateRef, watch, nextTick } from "vue";
import SlideStage from "./SlideStage.vue";
import AnimationStageLayer from "./AnimationStageLayer.vue";
import SlideTransitionLayer from "./SlideTransitionLayer.vue";
import SlideProofingLayer from "./SlideProofingLayer.vue";
import SpellContextMenu from "../SpellContextMenu.vue";
import UiButton from "../ui/UiButton.vue";
import { cmToPx } from "../../lib/core/units";

const props = defineProps({
  currentSlide: { type: Object, default: null },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  /** SPEC-094: the deck's theme, threaded to the stage surfaces. */
  theme: { type: Object, default: null },
  liveLabel: { type: String, default: "" },
  stageEdit: { type: Object, required: true },
  animations: { type: Object, required: true },
  transitions: { type: Object, required: true },
  /** SPEC-093: the current slide's pre-resolved deck field band. */
  fields: { type: Array, default: () => [] },
  /** SPEC-098: `{ search, spell }` — useDeckSearch and useSlideSpellCheck —
   * whose highlight and squiggle rects the proofing layer paints over the
   * stage, re-measured when the fit scale changes. */
  proofing: { type: Object, default: null },
});
const emit = defineEmits(["add"]);
const s = props.stageEdit;

const stageContainer = useTemplateRef("stageContainer");
const stageScale = ref(1);
const resizeObserver = typeof ResizeObserver !== "undefined" ? new ResizeObserver(recomputeScale) : null;

function recomputeScale() {
  const container = stageContainer.value;
  if (!container) {
    return;
  }
  const availableW = container.clientWidth;
  const availableH = container.clientHeight;
  const surfaceW = cmToPx(props.widthCm);
  const surfaceH = cmToPx(props.heightCm);
  if (surfaceW <= 0 || surfaceH <= 0 || availableW <= 0 || availableH <= 0) {
    return;
  }
  stageScale.value = Math.min(availableW / surfaceW, availableH / surfaceH);
}

watch(
  () => [props.currentSlide, props.widthCm, props.heightCm],
  () => void nextTick(recomputeScale),
);

watch(stageScale, () => {
  props.proofing?.search.remeasure();
  props.proofing?.spell.remeasure();
});

onMounted(() => {
  if (resizeObserver && stageContainer.value) {
    resizeObserver.observe(stageContainer.value);
  }
  window.addEventListener("resize", recomputeScale);
  void nextTick(recomputeScale);
});
onBeforeUnmount(() => {
  resizeObserver?.disconnect();
  window.removeEventListener("resize", recomputeScale);
});
</script>

<template>
  <!-- tabindex="-1": the find bar hands focus back here on close (SPEC-098). -->
  <div
    ref="stageContainer"
    class="stage-workspace"
    data-testid="stage-workspace"
    tabindex="-1"
  >
    <SlideStage
      v-if="currentSlide"
      :slide="currentSlide"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :scale="stageScale"
      :live-label="liveLabel"
      :fields="fields"
      :selected-frame-id="s.selection.selectedFrameId.value ?? undefined"
      :selected-frame-ids="s.selection.selectedFrameIds.value"
      :marquee="s.arrange.marquee.value ?? undefined"
      :rotating-deg="s.manipulation.rotatingDeg.value ?? undefined"
      :editing-frame-id="s.editing.editingFrameId.value ?? undefined"
      :editing-cell="s.editing.editingCell.value ?? undefined"
      :guides="s.manipulation.guides.value"
      :draw-preview="s.insertion.preview.value ?? undefined"
      :armed="s.insertion.isArmed.value"
      :theme="theme"
      @frame-select="s.onFrameSelect"
      @frame-start-edit="s.onStartEdit"
      @frame-start-edit-cell="s.onStartEditCell"
      @frame-cell-select="s.onCellSelect"
      @table-border-pointerdown="s.onTableBorderPointerDown"
      @frame-pointerdown="s.onFramePointerDown"
      @handle-pointerdown="s.onHandlePointerDown"
      @rotate-pointerdown="s.manipulation.onRotateHandlePointerDown"
      @adjust-pointerdown="s.manipulation.onAdjustPointerDown"
      @adjust-keydown="s.manipulation.onAdjustKeydown"
      @edit-mounted="s.editing.registerHost"
      @edit-unmounted="s.editing.registerHost(null)"
      @surface-pointerdown="s.onSurfaceDeselect"
      @surface-marquee="s.onSurfaceMarquee"
      @surface-draw="s.onSurfaceDraw"
      @surface-mounted="s.registerSurface"
    >
      <template
        v-if="proofing"
        #overlay
      >
        <SlideProofingLayer
          :highlight-rects="proofing.search.highlightRects.value"
          :underline-rects="proofing.spell.underlineRects.value"
          :scale="stageScale"
          @spell-context="proofing.spell.openContextMenuAt"
        />
      </template>
    </SlideStage>
    <div
      v-else
      class="empty-stage"
      data-testid="presentation-empty"
    >
      <h3>This deck has no slides</h3>
      <UiButton
        data-testid="presentation-empty-add"
        @click="emit('add')"
      >
        Add slide
      </UiButton>
    </div>
    <!-- SPEC-098: right-click on a squiggle (position: fixed, client px). -->
    <SpellContextMenu
      v-if="proofing && proofing.spell.contextMenu.value"
      :left="proofing.spell.contextMenu.value.left"
      :top="proofing.spell.contextMenu.value.top"
      :word="proofing.spell.contextMenu.value.word"
      :suggestions="proofing.spell.contextMenu.value.suggestions"
      @suggestion="proofing.spell.applySuggestion"
      @ignore="proofing.spell.ignore()"
      @ignore-all="proofing.spell.ignoreAll()"
      @add-to-dictionary="proofing.spell.addToDictionary()"
      @dismiss="proofing.spell.closeContextMenu()"
    />
    <AnimationStageLayer
      v-if="animations.open.value && currentSlide"
      :controller="animations"
      :slide="currentSlide"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :scale="stageScale"
      :theme="theme"
    />
    <SlideTransitionLayer
      v-if="transitions.previewState.value"
      class="stage-preview"
      :outgoing="transitions.previewState.value.outgoing"
      :incoming="transitions.previewState.value.incoming"
      :transition="transitions.previewState.value.transition"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :scale="stageScale"
      :theme="theme"
      @done="transitions.onPreviewDone()"
    />
  </div>
</template>

<style scoped>
.stage-workspace:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.stage-workspace {
  position: relative;
  flex: 1 1 auto;
  min-width: 0;
  min-height: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: var(--space-5);
  background: var(--surface-sunken);
  overflow: hidden;
}

.stage-preview {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 5;
}

.empty-stage {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-4);
  color: var(--text-secondary);
  text-align: center;
}
</style>
