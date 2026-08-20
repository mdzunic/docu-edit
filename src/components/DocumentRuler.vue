<script setup lang="ts">
import { computed, ref, useTemplateRef } from "vue";
import { markerPositions, valueFromMarkerPosition } from "../lib/ruler";
import { asTabStops } from "../lib/documentModel";

// The horizontal ruler above the page (SPEC-014 AC-22, AC-23): cm ticks
// across the page width, plus draggable markers for the current paragraph's
// start / first-line / end indents and the page margins. All geometry lives
// in lib/ruler.ts; this component only translates pointer x-coordinates to
// cm and emits the resulting value while a drag is in flight — the parent
// applies it through the command API, and the moved marker position flows
// back in through `formatting`, which is what makes the drag live.
const props = defineProps({
  pageWidthCm: { type: Number, required: true },
  marginLeftCm: { type: Number, required: true },
  marginRightCm: { type: Number, required: true },
  /** The first touched paragraph's formatting — indents plus direction,
   * which mirrors the ruler for RTL paragraphs (AC-25). */
  formatting: { type: Object, required: true },
  /** The caret section's column boundaries (SPEC-030 AC-14), or null when it
   * has a single column: `{ boundaries: [{ index, centerCm }] }` with each
   * gutter centre measured from the content area's start. */
  columns: { type: Object, default: null },
  /** The first touched paragraph's custom tab stops (SPEC-032 AC-24),
   * positions measured from the text area's start. */
  tabStops: { type: Array, default: () => [] },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["update", "open-tabs-dialog"]);

const root = useTemplateRef("root");

/** The marker kind being dragged, or "" while idle. */
const dragging = ref("");

/** The column boundary index being dragged, or -1 while idle. */
const columnDragging = ref(-1);

/** The section's column boundary markers (SPEC-030), positioned on the page:
 * a gutter centre is measured from the content area, so its page position adds
 * the left margin. */
const columnBoundaries = computed(() => {
  const columns = props.columns;
  if (!columns || !Array.isArray(columns.boundaries)) {
    return [];
  }
  return columns.boundaries.map((boundary = { index: 0, centerCm: 0 }) => ({
    index: boundary.index,
    pageCm: props.marginLeftCm + boundary.centerCm,
    centerCm: boundary.centerCm,
  }));
});

const context = computed(() => ({
  pageWidthCm: props.pageWidthCm,
  marginLeftCm: props.marginLeftCm,
  marginRightCm: props.marginRightCm,
}));

const positions = computed(() => markerPositions(context.value, props.formatting));

/** Whole-centimetre tick positions, skipping 0 and the page edge. */
const ticks = computed(() => {
  const marks = [];
  for (let cm = 1; cm < props.pageWidthCm; cm++) {
    marks.push(cm);
  }
  return marks;
});

const MARKERS = [
  { kind: "margin-left", position: "marginLeft", label: "Left margin" },
  { kind: "margin-right", position: "marginRight", label: "Right margin" },
  { kind: "start", position: "start", label: "Start indent" },
  { kind: "first-line", position: "firstLine", label: "First-line indent" },
  { kind: "end", position: "end", label: "End indent" },
];

function positionOf(marker = MARKERS[0]) {
  const values = positions.value;
  if (marker.position === "marginLeft") {
    return values.marginLeft;
  }
  if (marker.position === "marginRight") {
    return values.marginRight;
  }
  if (marker.position === "start") {
    return values.start;
  }
  if (marker.position === "firstLine") {
    return values.firstLine;
  }
  return values.end;
}

function positionCmFromEvent(event = new PointerEvent("pointermove")) {
  const el = root.value;
  if (!el) {
    return 0;
  }
  const rect = el.getBoundingClientRect();
  if (rect.width === 0) {
    return 0;
  }
  return ((event.clientX - rect.left) / rect.width) * props.pageWidthCm;
}

function startDrag(kind = "", event = new PointerEvent("pointerdown")) {
  if (props.disabled || !(event.target instanceof Element)) {
    return;
  }
  event.preventDefault();
  dragging.value = kind;
  event.target.setPointerCapture(event.pointerId);
}

function moveDrag(kind = "", event = new PointerEvent("pointermove")) {
  if (dragging.value !== kind) {
    return;
  }
  const update = valueFromMarkerPosition(
    // The template's plain string narrows through lib/ruler's own union via
    // this cast-free hop: MARKERS only ever holds valid kinds.
    kind === "start" || kind === "first-line" || kind === "end" || kind === "margin-left"
      ? kind
      : "margin-right",
    positionCmFromEvent(event),
    context.value,
    props.formatting,
  );
  emit("update", update);
}

function endDrag(kind = "", event = new PointerEvent("pointerup")) {
  if (dragging.value === kind) {
    dragging.value = "";
    if (event.target instanceof Element && event.target.hasPointerCapture(event.pointerId)) {
      event.target.releasePointerCapture(event.pointerId);
    }
  }
}

// --- custom tab stops (SPEC-032 AC-24) --------------------------------------

/** The alignment the tab-type selector currently arms; clicking the lane
 * lays a stop of this type. */
const tabType = ref("left");

// SPEC-054: the selector cycles all five alignments, including the new bar.
const TAB_TYPE_CYCLE = ["left", "center", "right", "decimal", "bar"];

const TAB_TYPE_GLYPHS = {
  left: "L",
  center: "⊥",
  right: "⌐",
  decimal: "⊥.",
  bar: "|",
};

const TAB_TYPE_LABELS = {
  left: "Left tab",
  center: "Centre tab",
  right: "Right tab",
  decimal: "Decimal tab",
  bar: "Bar tab",
};

function tabTypeGlyph(kind = "left") {
  if (kind === "center") {
    return TAB_TYPE_GLYPHS.center;
  }
  if (kind === "right") {
    return TAB_TYPE_GLYPHS.right;
  }
  if (kind === "decimal") {
    return TAB_TYPE_GLYPHS.decimal;
  }
  if (kind === "bar") {
    return TAB_TYPE_GLYPHS.bar;
  }
  return TAB_TYPE_GLYPHS.left;
}

function tabTypeLabel(kind = "left") {
  if (kind === "center") {
    return TAB_TYPE_LABELS.center;
  }
  if (kind === "right") {
    return TAB_TYPE_LABELS.right;
  }
  if (kind === "decimal") {
    return TAB_TYPE_LABELS.decimal;
  }
  if (kind === "bar") {
    return TAB_TYPE_LABELS.bar;
  }
  return TAB_TYPE_LABELS.left;
}

/** Click cycles the armed type; double-click opens the Tabs dialog. */
function cycleTabType() {
  if (props.disabled) {
    return;
  }
  const index = TAB_TYPE_CYCLE.indexOf(tabType.value);
  tabType.value = TAB_TYPE_CYCLE[(index + 1) % TAB_TYPE_CYCLE.length];
}

function openTabsDialog() {
  if (!props.disabled) {
    emit("open-tabs-dialog");
  }
}

/** Whether the current paragraph runs right-to-left: a tab stop's start-
 * relative position then mirrors onto the page's right (start) edge, exactly
 * as the indent markers do (SPEC-014, AC-35). */
const rtl = computed(() => props.formatting?.direction === "rtl");

/** A start-relative stop position mapped to its physical page cm. */
function pageCmForPosition(positionCm = 0) {
  return rtl.value
    ? props.pageWidthCm - props.marginRightCm - positionCm
    : props.marginLeftCm + positionCm;
}

/** The inverse of pageCmForPosition: a physical page cm back to a start-
 * relative position, snapped to the drag grid. */
function positionCmForPage(pageCm = 0) {
  const raw = rtl.value ? props.pageWidthCm - props.marginRightCm - pageCm : pageCm - props.marginLeftCm;
  return Math.round(raw / 0.05) * 0.05;
}

// --- tab handle drag (SPEC-054 AC-4/AC-5) ----------------------------------
// A handle drags live and, if released off the ruler, is removed. Its visual
// position follows the pointer through local drag state; the model commits
// only on release, so a mid-drag re-sort never invalidates the handle index.
const tabDragging = ref(-1);
const tabDragPageCm = ref(0);
const tabDragOff = ref(false);

/** The rendered handles: each stop positioned at its physical page cm, or the
 * live drag position for the handle being dragged. */
const tabHandles = computed(() =>
  asTabStops(props.tabStops).map((stop, index) => ({
    index,
    alignment: stop.alignment,
    positionCm: stop.positionCm,
    pageCm: index === tabDragging.value ? tabDragPageCm.value : pageCmForPosition(stop.positionCm),
    dragging: index === tabDragging.value,
    off: index === tabDragging.value && tabDragOff.value,
  })),
);

/** True when a pointer y is far enough outside the ruler band to count as
 * "dragged off" (SPEC-054 AC-5) — the removal gesture. */
function pointerOffRuler(event = new PointerEvent("pointermove")) {
  const el = root.value;
  if (!el) {
    return false;
  }
  const rect = el.getBoundingClientRect();
  return event.clientY < rect.top - 8 || event.clientY > rect.bottom + 8;
}

/** A click on the free lane lays a new stop of the armed type at the
 * clicked position (SPEC-032 AC-24), snapped like the drag markers. */
function handleTabLaneClick(event = new MouseEvent("click")) {
  if (props.disabled) {
    return;
  }
  const relativeCm = positionCmForPage(
    positionCmFromEvent(new PointerEvent("pointermove", { clientX: event.clientX })),
  );
  if (relativeCm <= 0) {
    return;
  }
  emit("update", { kind: "add-tab", positionCm: relativeCm, alignment: tabType.value });
}

function startTabDrag(index = 0, event = new PointerEvent("pointerdown")) {
  if (props.disabled || !(event.target instanceof Element)) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  tabDragging.value = index;
  tabDragPageCm.value = pageCmForPosition(asTabStops(props.tabStops)[index]?.positionCm ?? 0);
  tabDragOff.value = false;
  event.target.setPointerCapture(event.pointerId);
}

function moveTabDrag(index = 0, event = new PointerEvent("pointermove")) {
  if (tabDragging.value !== index) {
    return;
  }
  tabDragPageCm.value = positionCmFromEvent(event);
  tabDragOff.value = pointerOffRuler(event);
}

function endTabDrag(index = 0, event = new PointerEvent("pointerup")) {
  if (tabDragging.value !== index) {
    return;
  }
  const off = tabDragOff.value || pointerOffRuler(event);
  const positionCm = positionCmForPage(positionCmFromEvent(event));
  tabDragging.value = -1;
  tabDragOff.value = false;
  if (event.target instanceof Element && event.target.hasPointerCapture(event.pointerId)) {
    event.target.releasePointerCapture(event.pointerId);
  }
  if (props.disabled) {
    return;
  }
  if (off || positionCm <= 0) {
    // Dragged off the ruler (AC-5), or dragged past the start edge — remove.
    emit("update", { kind: "remove-tab", index });
  } else {
    emit("update", { kind: "move-tab", index, positionCm });
  }
}

// --- column boundary drag (SPEC-030 AC-14) ---------------------------------

function startColumnDrag(index = 0, event = new PointerEvent("pointerdown")) {
  if (props.disabled || !(event.target instanceof Element)) {
    return;
  }
  event.preventDefault();
  columnDragging.value = index;
  event.target.setPointerCapture(event.pointerId);
}

function moveColumnDrag(index = 0, event = new PointerEvent("pointermove")) {
  if (columnDragging.value !== index) {
    return;
  }
  const boundary = columnBoundaries.value.find((entry) => entry.index === index);
  if (!boundary) {
    return;
  }
  // The gutter centre moves with the boundary, so the drag delta is the
  // pointer's content-relative position minus the boundary's current centre.
  const contentCm = positionCmFromEvent(event) - props.marginLeftCm;
  emit("update", { kind: "column-boundary", boundary: index, deltaCm: contentCm - boundary.centerCm });
}

function endColumnDrag(index = 0, event = new PointerEvent("pointerup")) {
  if (columnDragging.value === index) {
    columnDragging.value = -1;
    if (event.target instanceof Element && event.target.hasPointerCapture(event.pointerId)) {
      event.target.releasePointerCapture(event.pointerId);
    }
  }
}
</script>

<template>
  <div
    ref="root"
    class="ruler print-exclude"
    data-testid="document-ruler"
    role="group"
    aria-label="Ruler"
    :style="{ width: `${pageWidthCm}cm` }"
  >
    <!-- Margin zones shaded so the printable region reads at a glance. -->
    <div
      class="margin-zone"
      :style="{ left: 0, width: `${marginLeftCm}cm` }"
    />
    <div
      class="margin-zone"
      :style="{ right: 0, width: `${marginRightCm}cm` }"
    />
    <div
      v-for="cm in ticks"
      :key="cm"
      class="tick"
      :style="{ left: `${cm}cm` }"
    >
      <span class="tick-label">{{ cm }}</span>
    </div>
    <!-- Custom tab stops (SPEC-032 AC-24): the free lane lays a stop of the
         armed type on click; clicking a laid handle removes it; the corner
         selector cycles the type (double-click opens the Tabs dialog). -->
    <div
      class="tab-lane"
      data-testid="ruler-tab-lane"
      :style="{ left: `${marginLeftCm}cm`, right: `${marginRightCm}cm` }"
      @click="handleTabLaneClick"
    />
    <button
      type="button"
      class="tab-type-selector"
      data-testid="ruler-tab-selector"
      :disabled="disabled"
      :title="`${tabTypeLabel(tabType)} — click to change, double-click for the Tabs dialog`"
      :aria-label="`Tab stop type: ${tabTypeLabel(tabType)}`"
      :data-tab-type="tabType"
      @click="cycleTabType"
      @dblclick="openTabsDialog"
    >
      {{ tabTypeGlyph(tabType) }}
    </button>
    <button
      v-for="handle in tabHandles"
      :key="`tab-${handle.index}`"
      type="button"
      class="tab-handle"
      :class="{ 'tab-handle-dragging': handle.dragging, 'tab-handle-off': handle.off }"
      :data-testid="`ruler-tab-stop-${handle.index}`"
      :data-tab-alignment="handle.alignment"
      :style="{ left: `${handle.pageCm}cm` }"
      :title="`${handle.alignment} tab stop — drag to move, drag off to remove`"
      :aria-label="`${handle.alignment} tab stop at ${handle.positionCm} cm`"
      :disabled="disabled"
      @pointerdown="startTabDrag(handle.index, $event)"
      @pointermove="moveTabDrag(handle.index, $event)"
      @pointerup="endTabDrag(handle.index, $event)"
      @pointercancel="endTabDrag(handle.index, $event)"
    >
      {{ tabTypeGlyph(handle.alignment) }}
    </button>
    <button
      v-for="marker in MARKERS"
      :key="marker.kind"
      type="button"
      class="marker"
      :class="`marker-${marker.kind}`"
      :data-testid="`ruler-marker-${marker.kind}`"
      :style="{ left: `${positionOf(marker)}cm` }"
      :title="marker.label"
      :aria-label="marker.label"
      :disabled="disabled"
      @pointerdown="startDrag(marker.kind, $event)"
      @pointermove="moveDrag(marker.kind, $event)"
      @pointerup="endDrag(marker.kind, $event)"
      @pointercancel="endDrag(marker.kind, $event)"
    />
    <!-- Column boundary handles (SPEC-030 AC-14): one per gutter of a
         multi-column section; dragging resizes the two adjacent columns. -->
    <button
      v-for="boundary in columnBoundaries"
      :key="`column-${boundary.index}`"
      type="button"
      class="marker marker-column-boundary"
      :data-testid="`ruler-column-boundary-${boundary.index}`"
      :style="{ left: `${boundary.pageCm}cm` }"
      title="Column boundary"
      aria-label="Column boundary"
      :disabled="disabled"
      @pointerdown="startColumnDrag(boundary.index, $event)"
      @pointermove="moveColumnDrag(boundary.index, $event)"
      @pointerup="endColumnDrag(boundary.index, $event)"
      @pointercancel="endColumnDrag(boundary.index, $event)"
    />
  </div>
</template>

<style scoped>
.ruler {
  position: relative;
  height: 1.5rem;
  margin: 0 auto 0.25rem;
  background: var(--surface-base);
  border: 1px solid var(--border-default);
  border-radius: 2px;
  box-sizing: border-box;
  user-select: none;
}

.margin-zone {
  position: absolute;
  top: 0;
  bottom: 0;
  background: var(--surface-canvas);
}

.tick {
  position: absolute;
  bottom: 0;
  width: 1px;
  height: 0.375rem;
  background: var(--border-strong);
}

.tick-label {
  position: absolute;
  bottom: 0.4rem;
  left: 0;
  transform: translateX(-50%);
  font-size: 0.5625rem;
  color: var(--text-secondary);
}

/* Every marker is a small draggable handle centred on its position. The
   shapes echo Word's ruler: down-pointing triangle for the first-line
   indent (top edge), up-pointing for start/end (bottom edge), thin bars
   for the margins. */
.marker {
  position: absolute;
  width: 0.625rem;
  height: 0.625rem;
  padding: 0;
  border: none;
  background: var(--accent);
  transform: translateX(-50%);
  cursor: ew-resize;
  touch-action: none;
}

.marker:disabled {
  cursor: default;
  opacity: 0.4;
}

.marker-first-line {
  top: 0;
  clip-path: polygon(0 0, 100% 0, 50% 100%);
}

.marker-start,
.marker-end {
  bottom: 0;
  clip-path: polygon(50% 0, 100% 100%, 0 100%);
}

.marker-margin-left,
.marker-margin-right {
  top: 0;
  bottom: 0;
  height: auto;
  width: 0.25rem;
  background: var(--border-strong);
  opacity: 0.7;
}

/* The tab lane (SPEC-032): the printable area's lower half, catching lane
   clicks under the indent markers (which sit above it in the DOM). */
.tab-lane {
  position: absolute;
  bottom: 0;
  height: 55%;
  cursor: copy;
}

/* The tab-type corner selector, at the ruler's left edge. */
.tab-type-selector {
  position: absolute;
  left: 2px;
  top: 50%;
  transform: translateY(-50%);
  width: 1rem;
  height: 1rem;
  padding: 0;
  font-size: 0.5625rem;
  line-height: 1;
  border: 1px solid var(--border-strong);
  border-radius: 2px;
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

/* A laid tab stop's handle in the lane. */
.tab-handle {
  position: absolute;
  bottom: 0;
  transform: translateX(-50%);
  width: 0.75rem;
  height: 0.75rem;
  padding: 0;
  border: none;
  background: transparent;
  color: var(--accent);
  font-size: 0.5625rem;
  font-weight: 700;
  line-height: 1;
  cursor: pointer;
}

.tab-type-selector:disabled,
.tab-handle:disabled {
  opacity: 0.4;
  cursor: default;
}

/* A handle being dragged tracks the pointer; when dragged off the ruler it
   fades to signal the release-to-remove gesture (SPEC-054 AC-4/AC-5). */
.tab-handle-dragging {
  cursor: grabbing;
}

.tab-handle-off {
  opacity: 0.3;
}

/* Column boundary handle: a full-height bar centred in the gutter. */
.marker-column-boundary {
  top: 0;
  bottom: 0;
  height: auto;
  width: 0.3125rem;
  background: var(--success);
  border-radius: 2px;
}
</style>
