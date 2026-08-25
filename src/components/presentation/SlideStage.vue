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
  /** SPEC-081 selection/editing overlays (the filmstrip thumbnails pass
   * neither, so they stay read-only). */
  selectedFrameId: { type: String, default: null },
  editingFrameId: { type: String, default: null },
});

const emit = defineEmits(["frame-select", "frame-start-edit", "edit-mounted", "edit-unmounted", "surface-pointerdown"]);

function onSurfacePointerdown() {
  // Frames stop propagation on their own pointer-down, so this only fires for a
  // click on empty surface — deselect (AC-4).
  emit("surface-pointerdown");
}

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
      @mousedown="onSurfacePointerdown"
    >
      <SlideFrameView
        v-for="frame in frames"
        :key="frame.id"
        :frame="frame"
        :resolve-media="resolveMedia"
        :selected="frame.id === selectedFrameId"
        :editing="frame.id === editingFrameId"
        @select="(id) => emit('frame-select', id)"
        @start-edit="(id, point) => emit('frame-start-edit', id, point)"
        @edit-mounted="(el) => emit('edit-mounted', el)"
        @edit-unmounted="() => emit('edit-unmounted')"
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
