<script setup lang="ts">
// One fit-scaled slide pane of the presenter view (SPEC-100): a read-only
// SlideStage letterboxed into whatever box the layout gives it, rescaled by a
// ResizeObserver (SlideStageArea's maths). The presenter's current-slide pane
// passes its own build engine's `frameStyles`; the next-slide pane passes
// none, so it renders fully built. With no slide it states why (`emptyText` —
// the end of the show). `blank` paints the pane solid when it is itself the
// show (single-window mode). Media never autoplays here: the presenter's panes
// are previews, and the audience window is the one that plays. Runtime
// declarations, not defineProps<{...}>() generics (espree lints .vue scripts).
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import SlideStage from "./SlideStage.vue";
import { cmToPx } from "../../lib/core/units";

const props = defineProps({
  slide: { type: Object, default: null },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  theme: { type: Object, default: null },
  fields: { type: Array, default: () => [] },
  frameStyles: { type: Object, default: () => ({}) },
  /** SPEC-105: the current pane's per-paragraph build styles. */
  paragraphStyles: { type: Object, default: () => ({}) },
  emptyText: { type: String, default: "" },
  emptyTestid: { type: String, default: undefined },
  /** "none" | "black" | "white" — SPEC-100's blank screen, when this pane is
   * the show. */
  blank: { type: String, default: "none" },
});

const container = useTemplateRef("container");
const scale = ref(1);
const resizeObserver = typeof ResizeObserver !== "undefined" ? new ResizeObserver(fit) : null;

function fit() {
  const box = container.value;
  if (!box) return;
  const surfaceW = cmToPx(props.widthCm);
  const surfaceH = cmToPx(props.heightCm);
  if (surfaceW <= 0 || surfaceH <= 0 || box.clientWidth <= 0 || box.clientHeight <= 0) return;
  scale.value = Math.min(box.clientWidth / surfaceW, box.clientHeight / surfaceH);
}

// The blank colours are the show's content, mandated solid #000/#fff exactly
// as the play surface paints them — not theme chrome.
const blankStyle = computed(() => ({ background: props.blank === "white" ? "#fff" : "#000" }));

watch(
  () => [props.slide, props.widthCm, props.heightCm],
  () => void nextTick(fit),
);
onMounted(() => {
  if (resizeObserver && container.value) resizeObserver.observe(container.value);
  window.addEventListener("resize", fit);
  void nextTick(fit);
});
onBeforeUnmount(() => {
  resizeObserver?.disconnect();
  window.removeEventListener("resize", fit);
});
</script>

<template>
  <div
    ref="container"
    class="slide-pane"
  >
    <SlideStage
      v-if="slide"
      :slide="slide"
      :width-cm="widthCm"
      :height-cm="heightCm"
      :media="media"
      :scale="scale"
      :interactive="false"
      :frame-styles="frameStyles"
      :paragraph-styles="paragraphStyles"
      :fields="fields"
      :theme="theme"
    />
    <p
      v-else-if="emptyText"
      class="pane-empty"
      :data-testid="emptyTestid"
    >
      {{ emptyText }}
    </p>
    <div
      v-if="blank !== 'none'"
      class="pane-blank"
      :style="blankStyle"
      aria-hidden="true"
    />
  </div>
</template>

<style scoped>
.slide-pane {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  min-width: 0;
  min-height: 0;
  overflow: hidden;
  border-radius: var(--radius-md);
  background: var(--surface-sunken);
}

.pane-empty {
  margin: 0;
  padding: var(--space-4);
  color: var(--text-secondary);
  font-size: var(--text-lg);
  text-align: center;
}

.pane-blank {
  position: absolute;
  inset: 0;
}
</style>
