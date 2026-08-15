<script setup lang="ts">
// The live rubber band while a shape is being drawn (SPEC-051 AC-3): the
// shape stretching with the pointer, plus a marquee around the box it will
// occupy. Painted with the very paintShape the page uses, from the very
// Shape useShapes is about to insert (drawPreview.shape) — so this is the
// shape you get, not a stand-in for it. Same canvas + DPR handling as
// DocShapeBlock.vue, whose comment explains the pattern.
//
// Transient and non-interactive: pointer-events stay off so the marquee can
// never swallow the mouseup that ends the very gesture drawing it.
//
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { onMounted, onUpdated, useTemplateRef } from "vue";
import { asShape } from "../lib/documentModel";
import { paintShapeIntoCanvas } from "../lib/shapeCanvas";

const props = defineProps({
  /** The documentModel.Shape this gesture will insert. */
  shape: { type: Object, required: true },
  /** The box it will occupy, page-relative px (same basis as .shape-handles). */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  width: { type: Number, required: true },
  height: { type: Number, required: true },
});

const canvasEl = useTemplateRef("canvasEl");

function repaint() {
  const canvas = canvasEl.value;
  if (!canvas) {
    return;
  }
  // A Line drags out with zero height; keep a paintable strip so its stroke
  // still has somewhere to land (DocShapeBlock clamps the same way).
  const widthPx = Math.max(1, Math.round(props.width));
  const heightPx = Math.max(1, Math.round(props.height));
  paintShapeIntoCanvas(canvas, asShape(props.shape), widthPx, heightPx);
}

onMounted(repaint);
onUpdated(repaint);

function boxStyle() {
  return {
    left: `${props.left}px`,
    top: `${props.top}px`,
    width: `${Math.max(1, props.width)}px`,
    height: `${Math.max(1, props.height)}px`,
  };
}
</script>

<template>
  <div
    class="shape-draw-preview"
    data-testid="shape-draw-preview"
    :style="boxStyle()"
  >
    <canvas
      ref="canvasEl"
      class="shape-draw-preview-canvas"
    />
  </div>
</template>

<style scoped>
.shape-draw-preview {
  position: absolute;
  /* Above the page's own content, below the selection handles — nothing
     here is interactive, so it only ever needs to be seen. */
  z-index: 20;
  pointer-events: none;
  outline: 1px dashed #1a73e8;
  outline-offset: 1px;
}

.shape-draw-preview-canvas {
  display: block;
  position: absolute;
  /* Half-drawn, so the shape reads as not-yet-committed while still showing
     its real fill and outline. */
  opacity: 0.7;
}
</style>
