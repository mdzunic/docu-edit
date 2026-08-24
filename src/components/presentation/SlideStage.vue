<script setup lang="ts">
// One slide rendered WYSIWYG at a given scale (SPEC-079): a fixed-size surface
// (cmToPx of the slide size) that CSS `transform: scale()` fits into its
// letterboxed viewport — the same fit-don't-reflow trick the editor's zoom
// uses. The surface is always rendered light (a deck's colours are its own):
// white unless the deck declares its own background. Runtime prop
// declarations, not defineProps<{...}>() generics (espree lints .vue scripts).
import { computed } from "vue";
import SlideFrameView from "./SlideFrameView.vue";
import { asMediaParts, asSlide, cmPx, mediaResolver, surfaceBackground } from "./slideRender";

const props = defineProps({
  /** The Slide to render (presentation model). */
  slide: { type: Object, required: true },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  /** The deck's media parts, for image frames. */
  media: { type: Array, default: () => [] },
  scale: { type: Number, default: 1 },
  /** When set, an aria-live region announces this label on change (AC-53).
   * Only the main stage passes it; filmstrip thumbnails do not. */
  liveLabel: { type: String, default: undefined },
});

const slide = computed(() => asSlide(props.slide));
const frames = computed(() => slide.value.frames);
const resolveMedia = computed(() => mediaResolver(asMediaParts(props.media)));

const viewportStyle = computed(() => ({
  width: cmPx(props.widthCm * props.scale),
  height: cmPx(props.heightCm * props.scale),
}));

const surfaceStyle = computed(() => ({
  width: cmPx(props.widthCm),
  height: cmPx(props.heightCm),
  transform: `scale(${props.scale})`,
  background: surfaceBackground(slide.value),
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
        v-for="frame in frames"
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
  box-shadow: var(--shadow-1);
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>
