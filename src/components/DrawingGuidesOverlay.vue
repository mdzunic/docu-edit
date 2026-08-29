<script setup lang="ts">
// Transient alignment guides drawn while a drawing is dragged, resized or
// created (SPEC-072). A screen-only overlay: it carries `print-exclude`, so
// print, print preview and PDF export omit it (App.vue / pdf/domExtract), and
// it never touches the model — the coordinates come from the gesture layer and
// vanish on release. Rendered as a `.pages` sibling in the same unscaled,
// `.pages`-relative pixel frame as ShapeHandlesOverlay, so a guide at a model
// centimetre lines up with the drawing it matched at any zoom.
import { computed } from "vue";
import { PX_PER_CM } from "../lib/core/units";
import { asSnapGuides } from "../lib/drawingSnap";

// Runtime prop declaration per the espree constraint (see AppShell.vue).
const props = defineProps({
  /** SnapGuide[] in `.pages`-relative centimetres: { axis, coordCm, startCm,
   * endCm }. axis "x" is a vertical line; "y" a horizontal line. */
  guides: { type: Array, default: () => [] },
});

// Each guide's absolute-positioned style, computed in the same unscaled
// `.pages` pixel frame the handles overlay uses. Typed via asSnapGuides so the
// map callback needs no annotation.
const guideLines = computed(() =>
  asSnapGuides(props.guides).map((guide) => {
    const coord = guide.coordCm * PX_PER_CM;
    const start = guide.startCm * PX_PER_CM;
    const length = (guide.endCm - guide.startCm) * PX_PER_CM;
    return guide.axis === "x"
      ? { left: `${coord}px`, top: `${start}px`, height: `${length}px`, width: "1px" }
      : { top: `${coord}px`, left: `${start}px`, width: `${length}px`, height: "1px" };
  }),
);
</script>

<template>
  <div
    class="drawing-guides print-exclude"
    data-testid="drawing-guides"
    aria-hidden="true"
  >
    <div
      v-for="(line, index) in guideLines"
      :key="index"
      class="drawing-guide"
      :style="line"
    />
  </div>
</template>

<style scoped>
.drawing-guides {
  position: absolute;
  inset: 0;
  z-index: 32;
  pointer-events: none;
}

.drawing-guide {
  position: absolute;
  background: var(--accent);
}
</style>
