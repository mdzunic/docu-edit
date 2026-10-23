<script setup lang="ts">
// The on-stage animation layer for the editor (SPEC-086): while the Animations
// panel is open it draws a numbered badge beside each animated frame (click →
// select that row, AC-15), and when the panel's "Play" runs it overlays a
// non-interactive preview that auto-plays the slide's builds on the editor
// stage (AC-11) using the same build engine as the show. The editor stage
// underneath always renders frames fully visible (AC-25) — only this overlay
// animates. Runtime prop declarations, not defineProps<{...}>() generics.
//
// SPEC-105: when the panel's "Edit path" arms the motion-path editor, this
// layer draws its overlay — the selected row's resolved path dashed from the
// frame's centre, a ghost of the frame at the path end, and a draggable
// endpoint (an imported custom path's refuses the drag). The geometry and the
// drag live in the controller's useMotionPathEdit; this only renders them.
import { computed, onBeforeUnmount, useTemplateRef, watch } from "vue";
import SlideStage from "./SlideStage.vue";
import { asSlide } from "./slideRender";
import { cmToPx } from "../../lib/core/units";
import { useBuildPreview } from "../../views/presentation/useBuildPreview";

const props = defineProps({
  /** The useAnimations() controller (badges, selectRow, previewState). */
  controller: { type: Object, required: true },
  slide: { type: Object, required: true },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  /** SPEC-094: the deck's theme, threaded to the stage surfaces. */
  theme: { type: Object, default: null },
  scale: { type: Number, default: 1 },
});

const c = props.controller;

/** A badge's pixel position at its frame's top-left corner (scaled). */
function badgeStyle(frameId = "") {
  const frame = asSlide(props.slide).frames.find((f) => f.id === frameId);
  if (!frame) return { display: "none" };
  return {
    left: `${cmToPx(frame.xCm) * props.scale}px`,
    top: `${cmToPx(frame.yCm) * props.scale}px`,
  };
}

const viewportStyle = computed(() => ({
  width: `${cmToPx(props.widthCm) * props.scale}px`,
  height: `${cmToPx(props.heightCm) * props.scale}px`,
}));

// --- Play preview -----------------------------------------------------------
const reducedMotion =
  typeof window !== "undefined" && typeof window.matchMedia === "function"
    ? window.matchMedia("(prefers-reduced-motion: reduce)").matches
    : false;
const buildPreview = useBuildPreview(() => ({ widthCm: props.widthCm, heightCm: props.heightCm }));
const previewing = computed(() => c.previewState.value !== null);

// The autoplay driver (timers, sequencing) lives in the composable so this
// .vue script stays annotation-free (espree). A new previewState starts a run.
watch(
  () => c.previewState.value,
  (slide) => {
    buildPreview.stop();
    if (slide) buildPreview.autoplay(asSlide(slide), reducedMotion, () => c.onPreviewDone());
  },
);

// --- motion-path overlay (SPEC-105) ---------------------------------------------
const pathOverlayEl = useTemplateRef("pathOverlayEl");
const pathOverlay = computed(() => (previewing.value ? null : c.pathEdit.overlay.value));
// The overlay draws in unscaled surface px; the endpoint keeps an 8px radius
// on screen whatever the fit scale, so it stays a target worth grabbing.
const endpointRadius = computed(() => 8 / (props.scale > 0 ? props.scale : 1));
function onEndpointDown(event = new PointerEvent("pointerdown")) {
  c.pathEdit.onEndpointPointerDown(event, pathOverlayEl.value);
}

function dismissPreview() {
  buildPreview.stop();
  c.onPreviewDone();
}

onBeforeUnmount(() => buildPreview.stop());
</script>

<template>
  <div class="anim-layer">
    <!-- Numbered badges over each animated frame -->
    <div
      v-if="!previewing"
      class="badge-viewport print-exclude"
      :style="viewportStyle"
    >
      <button
        v-for="badge in c.badges.value"
        :key="badge.index"
        type="button"
        class="badge"
        :style="badgeStyle(badge.frameId)"
        :data-testid="`animation-badge-${badge.index}`"
        :aria-label="`Animation ${badge.number}`"
        @click.stop="c.selectRow(badge.index)"
      >
        {{ badge.number }}
      </button>
    </div>

    <!-- Motion-path editor overlay (SPEC-105) -->
    <svg
      v-if="pathOverlay"
      ref="pathOverlayEl"
      class="path-overlay print-exclude"
      data-testid="animation-path-overlay"
      :style="viewportStyle"
      :viewBox="`0 0 ${pathOverlay.widthPx} ${pathOverlay.heightPx}`"
      role="group"
      aria-label="Motion path"
    >
      <rect
        class="path-ghost"
        data-testid="animation-path-ghost"
        :x="pathOverlay.ghost.x"
        :y="pathOverlay.ghost.y"
        :width="pathOverlay.ghost.width"
        :height="pathOverlay.ghost.height"
      />
      <path
        class="path-line"
        :d="pathOverlay.d"
      />
      <circle
        class="path-endpoint"
        :class="{ 'is-read-only': pathOverlay.readOnly }"
        data-testid="animation-path-endpoint"
        role="slider"
        tabindex="0"
        :aria-label="c.pathEdit.endpointLabel.value"
        :aria-disabled="pathOverlay.readOnly ? 'true' : 'false'"
        :cx="pathOverlay.handle.x"
        :cy="pathOverlay.handle.y"
        :r="endpointRadius"
        @pointerdown.stop="onEndpointDown"
        @pointermove="c.pathEdit.onEndpointPointerMove($event)"
        @pointerup="c.pathEdit.onEndpointPointerUp($event)"
        @pointercancel="c.pathEdit.onEndpointPointerUp($event)"
        @keydown="c.pathEdit.onEndpointKeydown($event)"
      />
    </svg>

    <!-- Non-interactive Play preview -->
    <div
      v-if="previewing"
      class="preview"
      data-testid="animation-preview"
      @click="dismissPreview"
    >
      <SlideStage
        :slide="slide"
        :width-cm="widthCm"
        :height-cm="heightCm"
        :media="media"
        :scale="scale"
        :interactive="false"
        :frame-styles="buildPreview.frameStyles.value"
        :paragraph-styles="buildPreview.paragraphStyles.value"
        :theme="theme"
      />
    </div>
  </div>
</template>

<style scoped>
.anim-layer {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  pointer-events: none;
}

.badge-viewport {
  position: relative;
}

.badge {
  position: absolute;
  transform: translate(-50%, -50%);
  min-width: 1.25rem;
  height: 1.25rem;
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-pill);
  background: var(--accent);
  color: var(--text-on-brand);
  font-size: var(--text-xs);
  line-height: 1;
  cursor: pointer;
  pointer-events: auto;
  font-variant-numeric: tabular-nums;
}

.path-overlay {
  position: absolute;
  overflow: visible;
  pointer-events: none;
}

.path-line {
  fill: none;
  stroke: var(--accent);
  stroke-width: 2;
  stroke-dasharray: 6 4;
  vector-effect: non-scaling-stroke;
}

.path-ghost {
  fill: none;
  stroke: var(--accent);
  stroke-width: 1;
  stroke-dasharray: 3 3;
  opacity: 0.6;
  vector-effect: non-scaling-stroke;
}

.path-endpoint {
  fill: var(--accent);
  stroke: var(--surface-default);
  stroke-width: 2;
  vector-effect: non-scaling-stroke;
  cursor: grab;
  pointer-events: auto;
  touch-action: none;
}

.path-endpoint.is-read-only {
  fill: var(--surface-default);
  stroke: var(--accent);
  cursor: not-allowed;
}

.preview {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-overlay);
  pointer-events: auto;
  z-index: 6;
}
</style>
