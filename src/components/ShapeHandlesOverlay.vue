<script setup lang="ts">
// SPEC-051: the selected shape's eight resize handles, its rotation handle and
// its per-geometry adjustment handles — the overlay factored out of
// DocumentView (unchanged markup, testids and classes) to keep the view within
// its line budget when SPEC-070 added the connector overlay beside it. Reacts
// to `geometry` (the measured handle box) changing so the style getters
// re-run. Runtime prop/emit declarations per the espree constraint.
import ShapeAdjustHandles from "./ShapeAdjustHandles.vue";
import { IMAGE_HANDLES } from "../lib/imageResize";

defineProps({
  /** The measured handle box (undefined hides the overlay); also the
   * reactivity trigger so `handleStyle`/`rotationStyle` re-run when it moves. */
  geometry: { type: Object, default: undefined },
  /** The per-geometry adjustment handles to draw. */
  adjustmentHandles: { type: Array, default: () => [] },
  /** `(handle) => styleObject` for a resize handle. */
  handleStyle: { type: Function, default: () => () => ({}) },
  /** `() => styleObject` for the rotation handle. */
  rotationStyle: { type: Function, default: () => () => ({}) },
});

defineEmits(["resize-start", "rotate-start", "adjust-start", "adjust-keydown"]);
</script>

<template>
  <div
    v-if="geometry"
    class="shape-handles print-exclude"
    data-testid="shape-handles"
  >
    <div
      v-for="handle in IMAGE_HANDLES"
      :key="handle"
      class="shape-resize-handle"
      :data-testid="`shape-resize-${handle}`"
      :style="handleStyle(handle)"
      @mousedown="$emit('resize-start', $event, handle)"
    />
    <div
      class="shape-rotation-handle"
      data-testid="shape-rotation-handle"
      :style="rotationStyle()"
      @mousedown="$emit('rotate-start', $event)"
    />
    <ShapeAdjustHandles
      :handles="adjustmentHandles"
      @adjust-start="(...args) => $emit('adjust-start', ...args)"
      @adjust-keydown="(...args) => $emit('adjust-keydown', ...args)"
    />
  </div>
</template>

<style scoped>
/* The selected shape's eight resize handles plus rotation handle (SPEC-051
   AC-10, AC-16) — same overlay convention as .image-handles. */
.shape-handles {
  position: absolute;
  inset: 0;
  /* Above .shape-toolbar (z-index: 30): any overlap the wrapping toolbar
     lands in (a shape near the viewport top, where it flips below) must
     leave the handles clickable, not merely visible — a covered rotation
     handle never receives its drag (AC-16, AC-38). */
  z-index: 31;
  pointer-events: none;
}

.shape-resize-handle,
.shape-rotation-handle {
  position: absolute;
  pointer-events: auto;
  width: 8px;
  height: 8px;
  border: 1px solid var(--accent);
  border-radius: 2px;
  background: var(--surface-base);
  box-shadow: 0 1px 2px rgb(0 0 0 / 0.2);
}

.shape-rotation-handle {
  border-radius: 50%;
  cursor: grab;
}

/* SPEC-068: shape handles enlarge for a coarse pointer so a fingertip can grab
   them (AC-42's touch-equivalent for handle drag). */
@media (pointer: coarse) {
  .shape-resize-handle,
  .shape-rotation-handle {
    width: 20px;
    height: 20px;
  }
}
</style>
