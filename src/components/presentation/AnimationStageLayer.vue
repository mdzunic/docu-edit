<script setup lang="ts">
// The on-stage animation layer for the editor (SPEC-086): while the Animations
// panel is open it draws a numbered badge beside each animated frame (click →
// select that row, AC-15), and when the panel's "Play" runs it overlays a
// non-interactive preview that auto-plays the slide's builds on the editor
// stage (AC-11) using the same build engine as the show. The editor stage
// underneath always renders frames fully visible (AC-25) — only this overlay
// animates. Runtime prop declarations, not defineProps<{...}>() generics.
import { computed, onBeforeUnmount, watch } from "vue";
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
const buildPreview = useBuildPreview();
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
      class="badge-viewport"
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
  padding: 0 0.25rem;
  border: var(--border-width) solid var(--accent, #3b82f6);
  border-radius: var(--radius-full, 999px);
  background: var(--accent, #3b82f6);
  color: #fff;
  font-size: var(--text-xs, 0.7rem);
  line-height: 1;
  cursor: pointer;
  pointer-events: auto;
  font-variant-numeric: tabular-nums;
}

.preview {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: color-mix(in srgb, #000 25%, transparent);
  pointer-events: auto;
  z-index: 6;
}
</style>
