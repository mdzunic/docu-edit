<script setup lang="ts">
import { inject, onMounted, onUpdated, ref, useTemplateRef, watch } from "vue";
import { asChart } from "../lib/documentModel";
import { chartDisplayStyle, drawingWrapSpacerStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { asTrackChangesDisplayMode } from "../lib/documentModel";
import { paintChart } from "../lib/core/chartCanvas";
import { cmToPx, pxToCm } from "../lib/core/units";
import { resolveThemeColor, useTheme } from "../composables/useTheme";

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
// SPEC-063 AC-20: the presentation-layer zoom, injected from DocumentView, so
// the backing store is painted at devicePixelRatio * zoom while the CSS size
// stays the model px — the .pages transform then scales it without blurring.
const documentZoom = inject("documentZoom", ref(1));

function repaint() {
  const chart = asChart(props.block);
  const canvas = canvasEl.value;
  if (!chart || !canvas) {
    return;
  }
  const widthPx = Math.max(1, Math.round(cmToPx(chart.widthCm)));
  const heightPx = Math.max(1, Math.round(cmToPx(Math.max(chart.heightCm, 0.05))));
  const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
  const scale = dpr * (documentZoom.value || 1);
  canvas.width = Math.max(1, Math.round(widthPx * scale));
  canvas.height = Math.max(1, Math.round(heightPx * scale));
  canvas.style.width = `${widthPx}px`;
  canvas.style.height = `${heightPx}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.clearRect(0, 0, widthPx, heightPx);
  // The chart's chrome (axis lines, gridlines, tick/label text) reads from the
  // theme tokens through resolveThemeColor (SPEC-059 AC-27/66/67); the series
  // colours stay document data and are never themed (AC-65). lib/** cannot
  // read CSS vars, so the view resolves them here and passes them in (AC-67).
  paintChart(ctx, chart, widthPx, heightPx, {
    axis: resolveThemeColor("--chart-axis"),
    gridline: resolveThemeColor("--chart-gridline"),
    text: resolveThemeColor("--chart-text"),
  });
}

// Reading the shared theme so a switch repaints the canvas with freshly
// resolved chrome colours (AC-28).
const { resolvedTheme } = useTheme();

onMounted(repaint);
onUpdated(repaint);
watch(() => props.block, repaint, { deep: true });
watch(resolvedTheme, repaint);
// Re-paint the backing store when the zoom changes so it stays crisp (AC-20).
watch(documentZoom, repaint);

const floatingWrap = ["square", "topAndBottom", "behind", "infront"];

// The text-column width the wrap spacer needs — same measured-in-component
// pattern as DocShapeBlock.vue.
const anchorEl = useTemplateRef("anchorEl");
const columnWidthCm = ref(0);

function measureColumnWidth() {
  const parent = anchorEl.value instanceof HTMLElement ? anchorEl.value.parentElement : null;
  const cm = parent ? pxToCm(parent.clientWidth) : 0;
  if (Math.abs(cm - columnWidthCm.value) > 0.01) {
    columnWidthCm.value = cm;
  }
}

onMounted(measureColumnWidth);
onUpdated(measureColumnWidth);

/** The in-flow float that reserves this chart's text room (square /
 * topAndBottom) — null for every other wrap mode. */
function wrapSpacerStyle() {
  return drawingWrapSpacerStyle(asChart(props.block), columnWidthCm.value);
}
</script>

<template>
  <span
    ref="anchorEl"
    class="doc-chart-anchor"
    :class="{ 'doc-chart-anchor-floating': floatingWrap.includes(block.wrap) }"
    :data-section="section"
    :data-block="blockIndex"
  >
    <!-- The wrap spacer — see DocShapeBlock.vue: the float the surrounding
         text yields to, making square/topAndBottom real. -->
    <span
      v-if="wrapSpacerStyle() && !trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-chart-wrap-spacer"
      data-testid="chart-wrap-spacer"
      aria-hidden="true"
      :style="wrapSpacerStyle()"
    />
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

/* A zero-height in-flow block (SPEC-078's zero-footprint floating anchor):
   hoistFloatingDrawingEntries renders it immediately before its anchor
   paragraph, so its top-left IS the paragraph's top-left — the origin the
   chart's offsets are measured from. A block (not inline-block, not
   position:absolute) for two reasons: the wrap spacer's float must escape
   into the page's text flow, and a block box is immune to floats, so the
   anchor stays pinned even when its own spacer would have pushed an inline
   anchor sideways (see DocShapeBlock.vue's identical rule). */
.doc-chart-anchor-floating {
  display: block;
  height: 0;
  position: relative;
}

.doc-chart-wrap-spacer {
  pointer-events: none;
}

.doc-chart {
  display: block;
  /* The canvas below is position:absolute; this box must contain it in every
     wrap mode — same fix as DocShapeBlock: an in-line-with-text chart's
     canvas otherwise escapes to the page's top-left corner. */
  position: relative;
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
  outline: 2px solid var(--accent);
  outline-offset: 1px;
}
</style>
