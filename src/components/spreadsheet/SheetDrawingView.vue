<script setup lang="ts">
// One sheet drawing's paint (SPEC-131 "The drawing layer"), filling a
// `width` × `height` box in 100%-zoom px (the grid's zoom transform scales
// it): what sheetDrawingLayout.ts' drawingPaint decides —
// - an image as an <img> (`sheet-drawing-image`) inside a clipping box, its
//   crop applied by scaling and offsetting it (croppedImageBox);
// - a shape through core/shapeCanvas' paintShapeIntoCanvas on a canvas
//   (`sheet-drawing-canvas`; the backing store at the zoom, the theme the
//   workbook's), an unsupported geometry painting its dashed placeholder, with
//   its DrawingText laid over in the DOM (`sheet-drawing-text`) — top, middle
//   or bottom, each paragraph aligned, the first run's size, colour and
//   weight;
// - a group member by member (`sheet-drawing-member`, each at its fractions
//   of the group's box, painted by this component in turn);
// - a chart (SPEC-132) through core/chartCanvas on a canvas
//   (`sheet-chart-canvas`, on the chart's white paper; the backing store at
//   the device pixel ratio — and, like the shape canvas, the zoom — the
//   theme the workbook's), repainted when its visual object, box, zoom or
//   theme changes; a chart outside the catalogue paints paintChart's own
//   placeholder. The canvas states what it plots for tests and tools:
//   `data-chart-type`, `data-chart-categories` and `data-chart-values` (JSON:
//   the categories, and each series' values with null for a gap),
//   `data-chart-legend`, and `data-chart-unsupported="true"` on the
//   placeholder;
// - and a face-less foreign drawing or an image the browser cannot show
//   (EMF, WMF, TIFF, a missing part) as the labelled placeholder
//   (`sheet-drawing-placeholder`): paintChart's unsupported-chart look, a
//   dashed 1.5 px #888888 rectangle with the label centred. A foreign
//   drawing with a face paints the face.
// `hostRotates`: the drawing's element applies the rotation (CSS, about the
// centre), so a top-level shape paints unrotated here; a group member keeps
// its own. It paints the user's document, so its styles are the document's
// own colours, not app tokens (tokens.spec.ts exempts it as it does
// SheetCellLayer). Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed, onMounted, useTemplateRef, watch } from "vue";
import { DEFAULT_CHART_CHROME, paintChart, paintChartIntoCanvas } from "../../lib/core/chartCanvas";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import { fontStack } from "../../lib/core/fontSubstitution";
import { asDocumentTheme, asDrawingContent, asImageSrc, croppedImageBox, drawingPaint, drawingTextView } from "../../views/spreadsheet/sheetDrawingLayout";

const props = defineProps({
  /** The drawing's SheetDrawingContent. */
  content: { type: Object, required: true },
  /** The box, 100%-zoom px. */
  width: { type: Number, default: 0 },
  height: { type: Number, default: 0 },
  /** The grid's zoom as a fraction: the canvas backing store's scale. */
  zoom: { type: Number, default: 1 },
  /** The workbook's DocumentTheme (absent: DEFAULT_THEME). */
  theme: { type: Object, default: null },
  /** mediaId → the image's data URI, "" when it cannot show. */
  imageSrc: { type: Function, default: undefined },
  /** The host element rotates the drawing (a top-level drawing). */
  hostRotates: { type: Boolean, default: false },
  /** SPEC-136: an extra class on a chart's canvas (a printed page's
   * `sheet-print-chart`, which the PDF extractor reads as an image). */
  chartCanvasClass: { type: String, default: "" },
});

const theme = computed(() => asDocumentTheme(props.theme));
const paint = computed(() => drawingPaint(asDrawingContent(props.content), asImageSrc(props.imageSrc)));
const imageBox = computed(() => {
  const p = paint.value;
  return croppedImageBox(p.kind === "image" ? p.crop : undefined, props.width, props.height);
});
const text = computed(() => {
  const p = paint.value;
  return p.kind === "shape" && p.text ? drawingTextView(p.text, theme.value) : null;
});
const textFont = fontStack("Calibri");

const canvas = useTemplateRef("canvas");

function repaint() {
  const element = canvas.value;
  const p = paint.value;
  if (!(element instanceof HTMLCanvasElement) || p.kind !== "shape") {
    return;
  }
  const visual = props.hostRotates ? { ...p.visual, rotationDeg: 0 } : p.visual;
  paintShapeIntoCanvas(element, visual, props.width, props.height, props.zoom, { theme: theme.value });
}

onMounted(repaint);
watch([paint, () => props.width, () => props.height, () => props.zoom, theme], repaint, { flush: "post" });

/** SPEC-132: a chart drawing's visual, null for any other kind. */
const chart = computed(() => {
  const content = asDrawingContent(props.content);
  return content.kind === "chart" ? content.visual : null;
});
const chartCategories = computed(() => JSON.stringify(chart.value?.categories ?? []));
const chartValues = computed(() => JSON.stringify((chart.value?.series ?? []).map((series) => series.values)));
const chartCanvas = useTemplateRef("chartCanvas");

function repaintChart() {
  const element = chartCanvas.value;
  const visual = chart.value;
  if (!(element instanceof HTMLCanvasElement) || !visual) {
    return;
  }
  const zoom = props.zoom || 1;
  if (zoom === 1) {
    paintChartIntoCanvas(element, visual, props.width, props.height, theme.value);
    return;
  }
  // paintChartIntoCanvas with the grid's zoom in the backing store, as the
  // shape canvas has it, so the zoom transform does not blur the chart.
  const width = Math.max(1, Math.ceil(props.width));
  const height = Math.max(1, Math.ceil(props.height));
  const scale = (window.devicePixelRatio || 1) * zoom;
  element.width = Math.max(1, Math.round(width * scale));
  element.height = Math.max(1, Math.round(height * scale));
  element.style.width = `${width}px`;
  element.style.height = `${height}px`;
  const ctx = element.getContext("2d");
  if (ctx) {
    ctx.setTransform(scale, 0, 0, scale, 0, 0);
    ctx.clearRect(0, 0, width, height);
    paintChart(ctx, visual, width, height, DEFAULT_CHART_CHROME, theme.value);
  }
}

onMounted(repaintChart);
watch([chart, () => props.width, () => props.height, () => props.zoom, theme], repaintChart, { flush: "post" });

function px(value = 0) {
  return `${value}px`;
}

/** A group member's box: its fractions of this one. */
function memberStyle(box = { x: 0, y: 0, w: 0, h: 0 }) {
  return { left: px(box.x * props.width), top: px(box.y * props.height), width: px(box.w * props.width), height: px(box.h * props.height) };
}
</script>

<template>
  <div
    class="sheet-drawing-paint"
    :style="{ width: px(props.width), height: px(props.height) }"
  >
    <canvas
      v-if="chart"
      ref="chartCanvas"
      class="sheet-drawing-canvas sheet-chart-canvas"
      :class="props.chartCanvasClass"
      data-testid="sheet-chart-canvas"
      :data-chart-type="chart.chartType"
      :data-chart-categories="chartCategories"
      :data-chart-values="chartValues"
      :data-chart-legend="chart.options.legend"
      :data-chart-unsupported="chart.unsupportedChartXml !== undefined ? 'true' : undefined"
      aria-hidden="true"
    />
    <div
      v-else-if="paint.kind === 'image'"
      class="sheet-drawing-crop"
    >
      <img
        class="sheet-drawing-img"
        data-testid="sheet-drawing-image"
        :src="paint.src"
        alt=""
        draggable="false"
        :style="{ left: px(imageBox.left), top: px(imageBox.top), width: px(imageBox.width), height: px(imageBox.height) }"
      >
    </div>
    <template v-else-if="paint.kind === 'shape'">
      <canvas
        ref="canvas"
        class="sheet-drawing-canvas"
        data-testid="sheet-drawing-canvas"
        aria-hidden="true"
      />
      <div
        v-if="text"
        class="sheet-drawing-text"
        data-testid="sheet-drawing-text"
        :style="{ ...text.style, justifyContent: text.justify, fontFamily: textFont }"
      >
        <p
          v-for="(paragraph, index) in text.paragraphs"
          :key="index"
          class="sheet-drawing-paragraph"
          :style="{ textAlign: paragraph.align }"
        >
          {{ paragraph.text }}
        </p>
      </div>
    </template>
    <template v-else-if="paint.kind === 'group'">
      <div
        v-for="(member, index) in paint.members"
        :key="index"
        class="sheet-drawing-member"
        data-testid="sheet-drawing-member"
        :style="memberStyle(member.box)"
      >
        <SheetDrawingView
          :content="member.content"
          :width="member.box.w * props.width"
          :height="member.box.h * props.height"
          :zoom="props.zoom"
          :theme="props.theme"
          :image-src="props.imageSrc"
        />
      </div>
    </template>
    <div
      v-else
      class="sheet-drawing-placeholder"
      data-testid="sheet-drawing-placeholder"
    >
      <span class="sheet-drawing-placeholder-label">{{ paint.label }}</span>
    </div>
  </div>
</template>

<style scoped>
/* The drawing's own paint: document colours, not app tokens (see the
   header). Overflow stays visible so a line's stroke and a rotated canvas's
   padding are not clipped; the image's crop box clips. */
.sheet-drawing-paint {
  position: relative;
}

.sheet-drawing-crop {
  position: absolute;
  inset: 0;
  overflow: hidden;
}

.sheet-drawing-img {
  position: absolute;
  max-width: none;
  user-select: none;
}

.sheet-drawing-canvas {
  position: absolute;
  pointer-events: none;
}

/* A chart's paper: Excel's and LibreOffice's default chart area is white. */
.sheet-chart-canvas {
  background: #ffffff;
}

/* Excel's default text insets: 0.1" left and right, 0.05" top and bottom. */
.sheet-drawing-text {
  position: absolute;
  inset: 0;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  padding: 4.8px 9.6px;
  overflow: hidden;
  line-height: 1.2;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.sheet-drawing-paragraph {
  margin: 0;
}

.sheet-drawing-member {
  position: absolute;
}

/* paintChart's unsupported-chart look. */
.sheet-drawing-placeholder {
  position: absolute;
  inset: 0;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  border: 1.5px dashed #888888;
  color: #666666;
  font: 10px sans-serif;
  text-align: center;
}

.sheet-drawing-placeholder-label {
  padding: 0 2px;
}
</style>
