<script setup lang="ts">
import { onMounted, onUpdated, useTemplateRef, watch } from "vue";
import { asBlock, asChart, asDocumentModel, asImage, asShape, isChart, isImage, isShape } from "../lib/documentModel";
import { paintShapeIntoCanvas } from "../lib/core/shapeCanvas";
import { paintChart } from "../lib/core/chartCanvas";
import { mediaDataUri } from "../lib/core/media";
import { cmToPx } from "../lib/core/units";
import { shapeTextRenderEntries } from "../lib/pageRender";
import { asListDefinitions } from "../lib/core/lists";
import DocBlockEntries from "./DocBlockEntries.vue";

// The painted visual of one leaf group member — a Shape, Chart or Image
// (SPEC-057 "Rendering and test surface"): the container composes these
// visuals with member proxies (DocDrawingGroupBlock.vue), so this component
// owns only the pixels, never a proxy of its own. It fills a wrapper the
// parent has already positioned and sized to the member's box, mirroring
// DocShapeBlock/DocChartBlock/DocImageBlock's own canvas painters without
// their anchor/proxy machinery. Runtime prop declarations per the espree
// constraint (see DocImageBlock.vue).
defineOptions({ inheritAttrs: false });

const props = defineProps({
  /** The member block (Shape | Chart | Image). */
  block: { type: Object, required: true },
  /** The document model — image members resolve their media part from it. */
  model: { type: Object, default: null },
  /** The document's multilevel list definitions — a text-box member's own
   * list paragraphs resolve markers/indents from the same catalogue. */
  listDefinitions: { type: Array, default: () => [] },
});

const canvasEl = useTemplateRef("canvasEl");

function repaint() {
  const block = asBlock(props.block);
  const canvas = canvasEl.value;
  if (!canvas) {
    return;
  }
  if (isShape(block)) {
    const shape = asShape(block);
    const widthPx = Math.max(1, Math.round(cmToPx(shape.widthCm)));
    const heightPx = Math.max(1, Math.round(cmToPx(Math.max(shape.heightCm, 0.05))));
    paintShapeIntoCanvas(canvas, shape, widthPx, heightPx);
    return;
  }
  if (isChart(block)) {
    const chart = asChart(block);
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
}

onMounted(repaint);
onUpdated(repaint);
watch(() => props.block, repaint, { deep: true });

function imageSrc() {
  const image = asImage(props.block);
  if (!props.model) {
    return "";
  }
  const parts = asDocumentModel(props.model).media ?? [];
  return mediaDataUri(parts.find((part) => part.mediaId === image.mediaId));
}

function isImageMember() {
  return isImage(asBlock(props.block));
}

function isShapeMember() {
  return isShape(asBlock(props.block));
}

/** A shape member's text frame as editable entries (AC-25): the container
 * shows it inside the member visual so a grouped text box can be entered. */
function shapeTextEntries() {
  const block = asBlock(props.block);
  if (!isShape(block)) {
    return [];
  }
  const shape = asShape(block);
  if (!shape.text) {
    return [];
  }
  return shapeTextRenderEntries(shape, asListDefinitions(props.listDefinitions));
}
</script>

<template>
  <img
    v-if="isImageMember()"
    class="doc-group-member-image"
    :src="imageSrc()"
    :alt="block.altText || ''"
    draggable="false"
  >
  <template v-else>
    <canvas
      ref="canvasEl"
      class="doc-group-member-canvas"
      aria-hidden="true"
    />
    <div
      v-if="isShapeMember() && block.text"
      class="doc-group-member-text"
      :class="`doc-shape-valign-${block.text.vAlign}`"
      data-canvas-ignore="1"
    >
      <DocBlockEntries :entries="shapeTextEntries()" />
    </div>
  </template>
</template>

<style scoped>
.doc-group-member-image {
  display: block;
  width: 100%;
  height: 100%;
  pointer-events: none;
}

.doc-group-member-canvas {
  display: block;
  position: absolute;
  top: 0;
  left: 0;
  pointer-events: none;
}

.doc-group-member-text {
  display: flex;
  flex-direction: column;
  position: absolute;
  inset: 0;
  padding: 0.1cm 0.2cm;
  overflow: hidden;
  white-space: pre-wrap;
  word-wrap: break-word;
  line-height: normal;
  pointer-events: none;
}

.doc-shape-valign-top {
  justify-content: flex-start;
}

.doc-shape-valign-middle {
  justify-content: center;
}

.doc-shape-valign-bottom {
  justify-content: flex-end;
}
</style>
