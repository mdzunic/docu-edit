<script setup lang="ts">
import { computed, ref, useTemplateRef } from "vue";
import {
  asRowBoundaries,
  valueFromVerticalPosition,
  verticalMarkerPositionCm,
  verticalRulerTicks,
  verticalTickScheme,
} from "../lib/verticalRuler";
import { asMeasurementUnit, CM_PER_PX, formatLength, fromCm, roundForDisplay } from "../lib/core/units";

// The vertical ruler down the left edge of the page in print layout (SPEC-064
// AC-1..6): unit-labelled ticks numbered outward from the text-area top, plus
// draggable markers for the top/bottom margins, the header/footer band
// boundaries, and the caret table's row heights. Like DocumentRuler.vue it is
// purely presentational — it emits a proposed `update` and the view turns it
// into an undoable model edit; the moved value flows back in through the props,
// which is what makes a drag live. It renders inside the scaled `.pages`
// subtree, so a ruler centimetre equals a document centimetre at every zoom
// (AC-1, AC-15) without the component doing any scaling of its own.
const props = defineProps({
  pageHeightCm: { type: Number, required: true },
  marginTopCm: { type: Number, required: true },
  marginBottomCm: { type: Number, required: true },
  /** Header band distance from the page's top edge (SPEC-029). */
  headerDistanceCm: { type: Number, default: 0 },
  /** Footer band distance from the page's bottom edge (SPEC-029). */
  footerDistanceCm: { type: Number, default: 0 },
  /** Row boundaries of the table containing the caret, each `{ row, topCm,
   * bottomCm }` measured from the page's top edge; empty when the caret is not
   * in a table (AC-5, AC-18). */
  tableRows: { type: Array, default: () => [] },
  /** "body" | "header" | "footer" — which band is being edited (AC-6). */
  activeBand: { type: String, default: "body" },
  /** The document's measurement-unit preference (AC-4, AC-9). */
  unit: { type: String, default: "cm" },
  disabled: { type: Boolean, default: false },
  /** SPEC-063 zoom fraction — used only to pick the tick density; the ruler is
   * scaled by its `.pages` ancestor, so its own geometry stays in cm (AC-1). */
  zoom: { type: Number, default: 1 },
  /** SPEC-063's client→content-pixel map (AC-15): a drag reads the pointer
   * through this rather than raw client coordinates, so it lands accurately at
   * any zoom. Passed from the view because it depends on the live `.pages`
   * rect. */
  clientToContentPx: {
    type: Function,
    default: (_clientX = 0, clientY = 0) => ({ x: 0, y: clientY }),
  },
});

const emit = defineEmits(["update"]);

const root = useTemplateRef("root");

/** The marker kind being dragged, or "" while idle. */
const dragging = ref("");
/** The table row index being dragged, or -1 while idle. */
const rowDragging = ref(-1);

const measurementUnit = computed(() => asMeasurementUnit(props.unit));

const context = computed(() => ({
  pageHeightCm: props.pageHeightCm,
  marginTopCm: props.marginTopCm,
  marginBottomCm: props.marginBottomCm,
  headerDistanceCm: props.headerDistanceCm,
  footerDistanceCm: props.footerDistanceCm,
}));

const ticks = computed(() =>
  verticalRulerTicks(props.pageHeightCm, props.marginTopCm, measurementUnit.value, props.zoom),
);

// The four band-boundary markers, each carrying the model value it edits (for
// its screen-reader read-out) and its physical position on the ruler.
const BAND_MARKERS = [
  { kind: "margin-top", label: "Top margin", value: "marginTop" },
  { kind: "margin-bottom", label: "Bottom margin", value: "marginBottom" },
  { kind: "header", label: "Header from top", value: "header" },
  { kind: "footer", label: "Footer from bottom", value: "footer" },
];

function bandKind(kind = "margin-top") {
  return kind === "margin-bottom" || kind === "header" || kind === "footer" ? kind : "margin-top";
}

function markerValueCm(which = "marginTop") {
  if (which === "marginTop") {
    return props.marginTopCm;
  }
  if (which === "marginBottom") {
    return props.marginBottomCm;
  }
  if (which === "header") {
    return props.headerDistanceCm;
  }
  return props.footerDistanceCm;
}

const markers = computed(() =>
  BAND_MARKERS.map((marker) => ({
    kind: marker.kind,
    label: marker.label,
    positionCm: verticalMarkerPositionCm(bandKind(marker.kind), context.value),
    valueCm: markerValueCm(marker.value),
  })),
);

/** The table row boundaries, each a draggable marker at its row's bottom edge
 * (AC-5); `topCm` lets a drag compute the row's new height. */
const rowMarkers = computed(() =>
  asRowBoundaries(props.tableRows).map((entry) => ({
    row: entry.row,
    topCm: entry.topCm,
    positionCm: entry.bottomCm,
    heightCm: Math.max(0, entry.bottomCm - entry.topCm),
  })),
);

const UNIT_NAMES = { cm: "centimetres", mm: "millimetres", in: "inches", pt: "points" };

function unitName() {
  const unit = measurementUnit.value;
  return unit === "mm" ? UNIT_NAMES.mm : unit === "in" ? UNIT_NAMES.in : unit === "pt" ? UNIT_NAMES.pt : UNIT_NAMES.cm;
}

/** The numeric value a slider announces via aria-valuenow, in the display
 * unit and rounded per roundForDisplay (AC-14a). */
function displayValue(valueCm = 0) {
  return roundForDisplay(fromCm(valueCm, measurementUnit.value), measurementUnit.value);
}

/** The spoken value, e.g. "Top margin, 2.5 centimetres" (AC-14a). */
function valueText(label = "", valueCm = 0) {
  return `${label}, ${formatLength(valueCm, measurementUnit.value)} ${unitName()}`;
}

/** The pointer's distance from the page top, in cm, read through SPEC-063's
 * client→content map so it lands accurately at any zoom (AC-15). */
function positionCmFromEvent(event = new PointerEvent("pointermove")) {
  return props.clientToContentPx(event.clientX, event.clientY).y * CM_PER_PX;
}

// --- band-boundary drag -----------------------------------------------------

function startDrag(kind = "", event = new PointerEvent("pointerdown")) {
  if (props.disabled || !(event.target instanceof Element)) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  dragging.value = kind;
  event.target.setPointerCapture(event.pointerId);
}

function moveDrag(kind = "", event = new PointerEvent("pointermove")) {
  if (dragging.value !== kind) {
    return;
  }
  emit("update", valueFromVerticalPosition(bandKind(kind), positionCmFromEvent(event), context.value));
}

function endDrag(kind = "", event = new PointerEvent("pointerup")) {
  if (dragging.value === kind) {
    dragging.value = "";
    if (event.target instanceof Element && event.target.hasPointerCapture(event.pointerId)) {
      event.target.releasePointerCapture(event.pointerId);
    }
  }
}

// --- table row-boundary drag (AC-5, AC-5a) ---------------------------------

function startRowDrag(index = 0, event = new PointerEvent("pointerdown")) {
  if (props.disabled || !(event.target instanceof Element)) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  rowDragging.value = index;
  event.target.setPointerCapture(event.pointerId);
}

function emitRowHeight(marker = { row: 0, topCm: 0 }, boundaryCm = 0) {
  emit("update", { kind: "table-row", row: marker.row, minHeightCm: Math.max(0, boundaryCm - marker.topCm) });
}

function moveRowDrag(index = 0, event = new PointerEvent("pointermove")) {
  if (rowDragging.value !== index) {
    return;
  }
  const marker = rowMarkers.value[index];
  if (marker) {
    emitRowHeight(marker, positionCmFromEvent(event));
  }
}

function endRowDrag(index = 0, event = new PointerEvent("pointerup")) {
  if (rowDragging.value === index) {
    rowDragging.value = -1;
    if (event.target instanceof Element && event.target.hasPointerCapture(event.pointerId)) {
      event.target.releasePointerCapture(event.pointerId);
    }
  }
}

// --- keyboard operation (AC-13) --------------------------------------------
// Arrow keys nudge a focused marker by one minor division, Shift+arrow by one
// major division (the same tick scheme the ruler draws). Up moves toward the
// page top, Down toward the bottom, matching the on-screen direction.

function nudgeStep(event = new KeyboardEvent("keydown")) {
  const scheme = verticalTickScheme(measurementUnit.value, props.zoom);
  const magnitude = event.shiftKey ? scheme.majorCm : scheme.minorCm;
  if (event.key === "ArrowUp" || event.key === "ArrowLeft") {
    return -magnitude;
  }
  if (event.key === "ArrowDown" || event.key === "ArrowRight") {
    return magnitude;
  }
  return 0;
}

function onMarkerKey(kind = "", positionCm = 0, event = new KeyboardEvent("keydown")) {
  if (props.disabled) {
    return;
  }
  const delta = nudgeStep(event);
  if (delta === 0) {
    return;
  }
  event.preventDefault();
  emit("update", valueFromVerticalPosition(bandKind(kind), positionCm + delta, context.value));
}

function onRowKey(index = 0, event = new KeyboardEvent("keydown")) {
  if (props.disabled) {
    return;
  }
  const delta = nudgeStep(event);
  const marker = rowMarkers.value[index];
  if (delta === 0 || !marker) {
    return;
  }
  event.preventDefault();
  emitRowHeight(marker, marker.positionCm + delta);
}
</script>

<template>
  <div
    ref="root"
    class="vertical-ruler print-exclude"
    data-testid="document-vertical-ruler"
    role="group"
    aria-label="Vertical ruler"
    :data-active-band="activeBand"
    :style="{ height: `${pageHeightCm}cm` }"
    @mousedown.stop
  >
    <!-- Margin zones shaded so the printable region reads at a glance. -->
    <div
      class="v-margin-zone"
      :style="{ top: 0, height: `${marginTopCm}cm` }"
    />
    <div
      class="v-margin-zone"
      :style="{ bottom: 0, height: `${marginBottomCm}cm` }"
    />
    <!-- The band being edited is highlighted so the focused region is obvious
         (AC-6): the header band spans [0, headerDistance] into the top margin;
         the footer band spans [footerDistance from bottom, bottom]. -->
    <div
      v-if="activeBand === 'header'"
      class="v-band-active"
      data-testid="vertical-ruler-active-band"
      :style="{ top: 0, height: `${Math.max(headerDistanceCm, marginTopCm)}cm` }"
    />
    <div
      v-if="activeBand === 'footer'"
      class="v-band-active"
      data-testid="vertical-ruler-active-band"
      :style="{ bottom: 0, height: `${Math.max(footerDistanceCm, marginBottomCm)}cm` }"
    />
    <div
      v-for="(tick, index) in ticks"
      :key="index"
      class="v-tick"
      :class="{ major: tick.major }"
      :style="{ top: `${tick.cm}cm` }"
    >
      <span
        v-if="tick.major"
        class="v-tick-label"
      >{{ tick.label }}</span>
    </div>
    <button
      v-for="marker in markers"
      :key="marker.kind"
      type="button"
      class="v-marker"
      :class="`v-marker-${marker.kind}`"
      :data-testid="`vertical-ruler-marker-${marker.kind}`"
      role="slider"
      :aria-label="marker.label"
      :aria-orientation="'vertical'"
      :aria-valuemin="0"
      :aria-valuemax="displayValue(pageHeightCm)"
      :aria-valuenow="displayValue(marker.valueCm)"
      :aria-valuetext="valueText(marker.label, marker.valueCm)"
      :title="marker.label"
      :style="{ top: `${marker.positionCm}cm` }"
      :disabled="disabled"
      @pointerdown="startDrag(marker.kind, $event)"
      @pointermove="moveDrag(marker.kind, $event)"
      @pointerup="endDrag(marker.kind, $event)"
      @pointercancel="endDrag(marker.kind, $event)"
      @keydown="onMarkerKey(marker.kind, marker.positionCm, $event)"
    />
    <button
      v-for="(marker, index) in rowMarkers"
      :key="`row-${marker.row}`"
      type="button"
      class="v-marker v-marker-row"
      :data-testid="`vertical-ruler-row-${marker.row}`"
      role="slider"
      :aria-label="`Table row ${marker.row + 1} height`"
      :aria-orientation="'vertical'"
      :aria-valuemin="0"
      :aria-valuemax="displayValue(pageHeightCm)"
      :aria-valuenow="displayValue(marker.heightCm)"
      :aria-valuetext="valueText(`Table row ${marker.row + 1} height`, marker.heightCm)"
      title="Table row height"
      :style="{ top: `${marker.positionCm}cm` }"
      :disabled="disabled"
      @pointerdown="startRowDrag(index, $event)"
      @pointermove="moveRowDrag(index, $event)"
      @pointerup="endRowDrag(index, $event)"
      @pointercancel="endRowDrag(index, $event)"
      @keydown="onRowKey(index, $event)"
    />
  </div>
</template>

<style scoped>
/* A fixed-width strip pinned to the page's left edge. It lives inside the
   scaled `.pages` subtree, so its cm-sized geometry scales with the document
   (a ruler cm == a document cm at every zoom) and it scrolls with the page. */
.vertical-ruler {
  position: absolute;
  top: 0;
  left: -0.9cm;
  width: 0.9cm;
  background: var(--surface-base);
  border: 1px solid var(--border-default);
  border-radius: 2px;
  box-sizing: border-box;
  user-select: none;
  z-index: 6;
}

.v-margin-zone {
  position: absolute;
  left: 0;
  right: 0;
  background: var(--surface-canvas);
}

/* The band under edit (AC-6): a translucent accent wash over the top or bottom
   band region so the focused header/footer is unmistakable. */
.v-band-active {
  position: absolute;
  left: 0;
  right: 0;
  background: color-mix(in srgb, var(--accent) 22%, transparent);
  pointer-events: none;
}

.v-tick {
  position: absolute;
  right: 0;
  width: 0.15cm;
  height: 0;
  border-top: 1px solid var(--border-strong);
}

.v-tick.major {
  width: 0.3cm;
}

.v-tick-label {
  position: absolute;
  right: 0.32cm;
  top: 50%;
  transform: translateY(-50%);
  font-size: 0.2cm;
  line-height: 1;
  color: var(--text-secondary);
  white-space: nowrap;
}

/* Every marker is a small draggable handle centred on its position, sitting on
   the ruler's inner (page-facing) edge. */
.v-marker {
  position: absolute;
  left: 0;
  right: 0;
  height: 0.12cm;
  padding: 0;
  border: none;
  background: var(--border-strong);
  transform: translateY(-50%);
  cursor: ns-resize;
  touch-action: none;
  opacity: 0.7;
}

.v-marker:disabled {
  cursor: default;
  opacity: 0.4;
}

.v-marker:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 1px;
}

.v-marker-header,
.v-marker-footer {
  background: var(--accent);
}

.v-marker-row {
  background: var(--success);
  opacity: 0.85;
}
</style>
