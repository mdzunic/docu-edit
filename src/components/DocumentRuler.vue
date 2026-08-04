<script setup lang="ts">
import { computed, ref, useTemplateRef } from "vue";
import { markerPositions, valueFromMarkerPosition } from "../lib/ruler";

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
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["update"]);

const root = useTemplateRef("root");

/** The marker kind being dragged, or "" while idle. */
const dragging = ref("");

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
</script>

<template>
  <div
    ref="root"
    class="ruler"
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
  </div>
</template>

<style scoped>
.ruler {
  position: relative;
  height: 1.5rem;
  margin: 0 auto 0.25rem;
  background: #fff;
  border: 1px solid #d0d0d0;
  border-radius: 2px;
  box-sizing: border-box;
  user-select: none;
}

.margin-zone {
  position: absolute;
  top: 0;
  bottom: 0;
  background: #f1f3f4;
}

.tick {
  position: absolute;
  bottom: 0;
  width: 1px;
  height: 0.375rem;
  background: #9aa0a6;
}

.tick-label {
  position: absolute;
  bottom: 0.4rem;
  left: 0;
  transform: translateX(-50%);
  font-size: 0.5625rem;
  color: #5f6368;
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
  background: #1a73e8;
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
  background: #80868b;
  opacity: 0.7;
}
</style>
