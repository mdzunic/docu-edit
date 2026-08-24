<script setup lang="ts">
// One slide rendered WYSIWYG at a given scale (SPEC-079): a fixed-size surface
// (cmToPx of the slide size) that CSS `transform: scale()` fits into its
// letterboxed viewport — the same fit-don't-reflow trick the editor's zoom
// uses. The surface is always rendered light (a deck's colours are its own):
// white unless the deck declares its own background. Dumb: props in, no store.
import { computed } from "vue";
import { cmToPx } from "../../lib/core/units";
import { mediaDataUri, type MediaPart } from "../../lib/core/media";
import type { Slide } from "../../lib/presentation/model";
import SlideFrameView from "./SlideFrameView.vue";

const props = withDefaults(
  defineProps<{
    slide: Slide;
    widthCm: number;
    heightCm: number;
    media?: MediaPart[];
    scale?: number;
    /** When set, an aria-live region announces this label on change (AC-53).
     * Only the main stage passes it; filmstrip thumbnails do not. */
    liveLabel?: string;
  }>(),
  { media: () => [], scale: 1 },
);

function px(cm: number): string {
  return `${cmToPx(cm)}px`;
}

const mediaMap = computed(() => new Map((props.media ?? []).map((part) => [part.mediaId, part])));

function resolveMedia(mediaId: string): string {
  return mediaDataUri(mediaMap.value.get(mediaId));
}

const surfaceBackground = computed(() => {
  const bg = props.slide.background;
  return bg && bg.type === "solid" ? bg.color : "#ffffff";
});

const viewportStyle = computed(() => ({
  width: px(props.widthCm * (props.scale ?? 1)),
  height: px(props.heightCm * (props.scale ?? 1)),
}));

const surfaceStyle = computed(() => ({
  width: px(props.widthCm),
  height: px(props.heightCm),
  transform: `scale(${props.scale ?? 1})`,
  background: surfaceBackground.value,
}));
</script>

<template>
  <div
    class="stage-viewport"
    :style="viewportStyle"
    data-testid="slide-stage"
  >
    <div
      class="stage-surface"
      :style="surfaceStyle"
    >
      <SlideFrameView
        v-for="frame in slide.frames"
        :key="frame.id"
        :frame="frame"
        :resolve-media="resolveMedia"
      />
    </div>
    <div
      v-if="liveLabel !== undefined"
      class="sr-only"
      aria-live="polite"
      data-testid="slide-live"
    >
      {{ liveLabel }}
    </div>
  </div>
</template>

<style scoped>
.stage-viewport {
  position: relative;
  overflow: hidden;
}

.stage-surface {
  position: absolute;
  top: 0;
  left: 0;
  transform-origin: top left;
  box-shadow: var(--shadow-1, 0 1px 3px rgba(0, 0, 0, 0.2));
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
