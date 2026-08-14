<script setup lang="ts">
import { onMounted, onUpdated, useTemplateRef, watch } from "vue";
import { asChart } from "../lib/documentModel";
import { chartDisplayStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { asTrackChangesDisplayMode } from "../lib/documentModel";
import { paintChart } from "../lib/chartCanvas";
import { cmToPx } from "../lib/images";

// One chart block on the page (SPEC-052). Runtime prop declarations, not
// defineProps<{...}>() generics — see DocImageBlock.vue's comment on the
// same espree constraint. Paints title/legend/axes/series on a canvas
// (CLAUDE.md's "HTML Canvas for complex shapes") and owns one absolutely
// positioned, transparent DOM proxy over its bounds — the element every
// e2e assertion, keyboard action and accessibility scan goes through,
// mirroring DocShapeBlock.vue's own canvas+proxy split (SPEC-051
// "Rendering and test surface"). A chart has no editable text frame (unlike
// a shape), so the proxy is always role="img" with no interior DOM.
defineOptions({ inheritAttrs: false });

const props = defineProps({
  /** The Chart block (documentModel.Chart). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  selected: { type: Boolean, default: false },
  /** This drawing's 0-based index in document order — data-testid="drawing-<n>",
   * shared with Shape's own counter (SPEC-051). */
  drawingIndex: { type: Number, default: 0 },
  displayMode: { type: String, default: "all" },
});

defineEmits(["select", "edit-data"]);

const canvasEl = useTemplateRef("canvasEl");

function repaint() {
  const chart = asChart(props.block);
  const canvas = canvasEl.value;
  if (!chart || !canvas) {
    return;
  }
  const widthPx = Math.max(1, Math.round(cmToPx(chart.widthCm)));
  const heightPx = Math.max(1, Math.round(cmToPx(Math.max(chart.heightCm, 0.05))));
  const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
  canvas.width = widthPx * dpr;
  canvas.height = heightPx * dpr;
  canvas.style.width = `${widthPx}px`;
  canvas.style.height = `${heightPx}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, widthPx, heightPx);
  paintChart(ctx, chart, widthPx, heightPx);
}

onMounted(repaint);
onUpdated(repaint);
watch(() => props.block, repaint, { deep: true });

const floatingWrap = ["square", "topAndBottom", "behind", "infront"];
</script>

<template>
  <span
    class="doc-chart-anchor"
    :class="{ 'doc-chart-anchor-floating': floatingWrap.includes(block.wrap) }"
    :data-section="section"
    :data-block="blockIndex"
  >
    <span
      v-show="!trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-chart"
      :class="{ 'doc-chart-selected': selected }"
      :style="chartDisplayStyle(asChart(block))"
    >
      <canvas
        ref="canvasEl"
        class="doc-chart-canvas"
        aria-hidden="true"
      />
      <span
        class="doc-chart-proxy"
        :data-testid="`drawing-${drawingIndex}`"
        role="img"
        :aria-label="block.altText || (block.unsupportedChartXml !== undefined ? 'Chart not supported' : 'Chart')"
        tabindex="0"
        @click.stop="(event) => $emit('select', event.shiftKey)"
        @dblclick.stop="$emit('edit-data')"
      />
    </span>
  </span>
</template>

<style scoped>
.doc-chart-anchor {
  display: inline-block;
  line-height: 0;
}

.doc-chart-anchor-floating {
  position: relative;
}

.doc-chart {
  display: block;
  cursor: pointer;
}

.doc-chart-canvas {
  display: block;
  position: absolute;
  top: 0;
  left: 0;
  pointer-events: none;
}

.doc-chart-proxy {
  display: block;
  position: relative;
  width: 100%;
  height: 100%;
  background: transparent;
  outline: none;
}

.doc-chart-selected .doc-chart-proxy {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
</style>
