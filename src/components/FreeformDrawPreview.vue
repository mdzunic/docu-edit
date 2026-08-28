<script setup lang="ts">
// SPEC-071: the live freeform/curve/scribble draw preview — a dashed SVG
// outline drawn over the pages in the same unscaled px space
// ShapeDrawPreview uses. `d` is a ready-made SVG path string in page px
// (useShapes.freeformPreviewD); this component only paints it. Runtime prop
// declaration per the espree constraint the .vue linter enforces.
defineProps({ d: { type: String, required: true } });
</script>

<template>
  <svg
    class="freeform-draw-preview print-exclude"
    data-testid="freeform-draw-preview"
    aria-hidden="true"
  >
    <path :d="d" />
  </svg>
</template>

<style scoped>
/* Overflow is visible so a stroke running past the nominal viewport shows. */
.freeform-draw-preview {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  overflow: visible;
  pointer-events: none;
  z-index: 5;
}

.freeform-draw-preview path {
  fill: none;
  stroke: var(--accent, #2e5395);
  stroke-width: 1.5;
  stroke-dasharray: 4 3;
}
</style>
