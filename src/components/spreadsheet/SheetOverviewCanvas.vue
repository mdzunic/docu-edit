<script setup lang="ts">
// SPEC-137's overview rendering: the canvas that paints the sheet — fills,
// borders, gridlines, headers and text (greeked below GRID_GREEK_BELOW_PX) —
// when the viewport alone would pass the grid's ceiling (sheetRenderPlan).
// It sits in the grid's scaled content box over the paper (whose one
// remaining element is the active cell) and under the overlays, covering the
// headers too; sheetOverview's painter sizes its backing store to device
// pixels and repaints at most once per animation frame. It takes no pointer:
// presses reach the grid, which hit-tests through GridMetrics. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { onBeforeUnmount, onMounted, useTemplateRef } from "vue";
import { asOverviewFrame, mountOverviewPainter } from "../../views/spreadsheet/sheetOverview";

const props = defineProps({
  /** () => OverviewFrame: what to paint, read inside the painter's effect. */
  frame: { type: Function, required: true },
});

const canvas = useTemplateRef("canvas");
let dispose = () => {};

onMounted(() => {
  if (canvas.value) {
    dispose = mountOverviewPainter(canvas.value, asOverviewFrame(props.frame));
  }
});
onBeforeUnmount(() => dispose());
</script>

<template>
  <canvas
    ref="canvas"
    class="sheet-overview-canvas"
    data-testid="sheet-overview-canvas"
    aria-hidden="true"
  />
</template>

<style scoped>
.sheet-overview-canvas {
  position: absolute;
  top: 0;
  left: 0;
  pointer-events: none;
}
</style>
