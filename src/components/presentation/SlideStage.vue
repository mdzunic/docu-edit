<script setup lang="ts">
// One slide rendered WYSIWYG at a given scale (SPEC-079): a fixed-size surface
// (cmToPx of the slide size) that CSS `transform: scale()` fits into its
// letterboxed viewport — the same fit-don't-reflow trick the editor's zoom
// uses. SPEC-082 adds direct manipulation over the surface: a draw-gesture
// preview, the selection overlay (resize handles + alignment guides), and
// pointer forwarding for move/resize. The surface is always rendered light (a
// deck's colours are its own): white unless the deck declares its own
// background. Runtime prop declarations, not defineProps<{...}>() generics
// (espree lints .vue scripts).
import { computed, onBeforeUnmount, onMounted, useTemplateRef } from "vue";
import SlideFrameView from "./SlideFrameView.vue";
import FrameSelectionOverlay from "./FrameSelectionOverlay.vue";
import { asMediaParts, asSlide, asStageFields, cmPx, mediaResolver, surfaceBackground } from "./slideRender";

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
  /** SPEC-091: the cell of the editing table currently in a session, so its
   * `<td>` renders the in-cell contenteditable host. */
  editingCell: { type: Object, default: null },
  /** SPEC-082 direct-manipulation surface state. */
  guides: { type: Object, default: () => ({ vertical: [], horizontal: [] }) },
  drawPreview: { type: Object, default: null },
  /** Whether an insertion tool is armed (the surface shows a crosshair). */
  armed: { type: Boolean, default: false },
  /** False for the filmstrip thumbnails: their frames render read-only, with
   * no selection chrome, focus, test hooks or pointer handling. */
  interactive: { type: Boolean, default: true },
  /** SPEC-086 playback: per-frame inline styles (by frame id) the build engine
   * layers on during a show/preview. Absent in the editor (frames fully
   * visible, AC-25). */
  frameStyles: { type: Object, default: () => ({}) },
  /** SPEC-093: the slide's pre-resolved deck fields (footer/slide-number/date),
   * computed by the parent from `fieldText` so this shared surface stays
   * framework-neutral and updates under the deck's shallowRef commit. Empty on a
   * suppressed title slide. */
  fields: { type: Array, default: () => [] },
});

const emit = defineEmits([
  "frame-select",
  "frame-start-edit",
  "frame-start-edit-cell",
  "frame-cell-select",
  "table-border-pointerdown",
  "frame-pointerdown",
  "handle-pointerdown",
  "edit-mounted",
  "edit-unmounted",
  "surface-pointerdown",
  "surface-draw",
  "surface-mounted",
]);

function onSurfaceMousedown() {
  // Frames stop propagation on their own pointer-down, so this only fires for a
  // click on empty surface — deselect (AC-4). Draw gestures are handled on
  // pointerdown (onSurfacePointerdown) before this fires.
  emit("surface-pointerdown");
}

function onSurfacePointerdown(event = new PointerEvent("pointerdown")) {
  // While a tool is armed, a pointerdown on the empty surface starts a draw
  // gesture; the parent stops propagation so the mousedown deselect is skipped.
  if (props.armed) {
    emit("surface-draw", event);
  }
}

const surfaceEl = useTemplateRef("surfaceEl");
onMounted(() => {
  if (surfaceEl.value) emit("surface-mounted", surfaceEl.value);
});
onBeforeUnmount(() => emit("surface-mounted", null));

const slide = computed(() => asSlide(props.slide));
const frames = computed(() => slide.value.frames);
const stageFieldList = computed(() => asStageFields(props.fields));
const resolveMedia = computed(() => mediaResolver(asMediaParts(props.media)));
const selectedFrame = computed(() => frames.value.find((frame) => frame.id === props.selectedFrameId) ?? null);
const showOverlay = computed(() => selectedFrame.value !== null && props.editingFrameId !== selectedFrame.value?.id);

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

const previewStyle = computed(() => {
  const p = props.drawPreview;
  if (!p) return null;
  return { left: cmPx(p.xCm), top: cmPx(p.yCm), width: cmPx(p.widthCm), height: cmPx(p.heightCm) };
});
</script>

<template>
  <div
    class="stage-viewport"
    :style="viewportStyle"
    data-testid="slide-stage"
  >
    <div
      ref="surfaceEl"
      class="stage-surface"
      :class="{ 'is-armed': armed }"
      :style="surfaceStyle"
      @mousedown="onSurfaceMousedown"
      @pointerdown="onSurfacePointerdown"
    >
      <SlideFrameView
        v-for="frame in frames"
        :key="frame.id"
        :frame="frame"
        :resolve-media="resolveMedia"
        :selected="frame.id === selectedFrameId"
        :editing="frame.id === editingFrameId"
        :editing-cell="frame.id === editingFrameId ? editingCell : undefined"
        :armed="armed"
        :top-level="interactive"
        :play-style="frameStyles[frame.id] || null"
        @select="(id) => emit('frame-select', id)"
        @start-edit="(id, point) => emit('frame-start-edit', id, point)"
        @start-edit-cell="(id, row, col, point) => emit('frame-start-edit-cell', id, row, col, point)"
        @cell-select="(id, row, col) => emit('frame-cell-select', id, row, col)"
        @table-border-pointerdown="(id, axis, index, event) => emit('table-border-pointerdown', id, axis, index, event)"
        @frame-pointerdown="(id, event) => emit('frame-pointerdown', id, event)"
        @edit-mounted="(el) => emit('edit-mounted', el)"
        @edit-unmounted="() => emit('edit-unmounted')"
      />
      <div
        v-for="(field, index) in stageFieldList"
        :key="`slide-field-${index}`"
        class="slide-field"
        data-testid="slide-field"
        :data-field-kind="field.kindToken"
        :style="field.style"
      >{{ field.text }}</div>
      <FrameSelectionOverlay
        v-if="showOverlay && selectedFrame"
        :frame="selectedFrame"
        :guides="guides"
        :width-cm="widthCm"
        :height-cm="heightCm"
        @handle-pointerdown="(handle, event) => emit('handle-pointerdown', selectedFrameId, handle, event)"
        @edit-request="(point) => emit('frame-start-edit', selectedFrameId, point)"
      />
      <div
        v-if="previewStyle"
        class="draw-preview"
        aria-hidden="true"
        :style="previewStyle"
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

.stage-surface.is-armed {
  cursor: crosshair;
}

.draw-preview {
  position: absolute;
  border: var(--border-width) dashed var(--accent);
  background: color-mix(in srgb, var(--accent) 12%, transparent);
  pointer-events: none;
}

/* SPEC-093 deck field band: a positioned, non-interactive label. Horizontal
 * alignment rides `justify-content` in the inline style (per the field). */
.slide-field {
  position: absolute;
  display: flex;
  align-items: center;
  overflow: hidden;
  white-space: nowrap;
  pointer-events: none;
  line-height: 1.1;
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
