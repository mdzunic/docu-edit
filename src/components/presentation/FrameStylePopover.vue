<script setup lang="ts">
// The style popover for the selected frame (SPEC-082): fill colour (with "No
// fill") and stroke colour/width for shapes and text boxes; for images, stroke
// only. Each committed change is one undo entry — the parent records + persists
// on every `set-fill`/`set-stroke`. Runtime prop/emit declarations per the
// espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSlideFrame } from "./slideRender";
import { MAX_STROKE_WIDTH_PT, MIN_STROKE_WIDTH_PT } from "../../lib/presentation/frames";

const props = defineProps({
  /** The selected SlideFrame (shape / text / image). */
  frame: { type: Object, required: true },
});

const emit = defineEmits(["set-fill", "set-stroke"]);

const frame = computed(() => asSlideFrame(props.frame));
const isImage = computed(() => frame.value.kind === "image");
const canFill = computed(() => frame.value.kind === "shape" || frame.value.kind === "text");

const fillColor = computed(() => {
  const f = frame.value;
  return (f.kind === "shape" || f.kind === "text") && f.fill?.type === "solid" ? f.fill.color : "#4472c4";
});
const hasFill = computed(() => {
  const f = frame.value;
  return (f.kind === "shape" || f.kind === "text") && f.fill?.type === "solid";
});

const stroke = computed(() => {
  const f = frame.value;
  return f.kind === "shape" || f.kind === "text" || f.kind === "image" ? f.stroke : undefined;
});
const strokeColor = computed(() => stroke.value?.color ?? "#000000");
const strokeWidth = computed(() => stroke.value?.widthPt ?? 1);
const hasStroke = computed(() => stroke.value !== undefined);

function inputValue(event = new Event("input")) {
  const target = event.target;
  return target instanceof HTMLInputElement ? target.value : "";
}
function setFillColor(event = new Event("input")) {
  emit("set-fill", { type: "solid", color: inputValue(event) });
}
function noFill() {
  emit("set-fill", { type: "none" });
}
function setStrokeColor(event = new Event("input")) {
  emit("set-stroke", { color: inputValue(event), widthPt: strokeWidth.value });
}
function setStrokeWidth(event = new Event("change")) {
  emit("set-stroke", { color: strokeColor.value, widthPt: Number(inputValue(event)) });
}
function toggleStroke() {
  emit("set-stroke", hasStroke.value ? null : { color: strokeColor.value, widthPt: strokeWidth.value });
}
</script>

<template>
  <div
    class="style-popover"
    role="group"
    aria-label="Frame style"
    data-testid="frame-style-popover"
  >
    <div
      v-if="canFill"
      class="row"
    >
      <span class="row-label">Fill</span>
      <input
        type="color"
        :value="fillColor"
        aria-label="Fill colour"
        data-testid="style-fill-color"
        @change="setFillColor"
      >
      <button
        type="button"
        class="pill"
        :aria-pressed="!hasFill"
        data-testid="style-no-fill"
        @click="noFill"
      >
        No fill
      </button>
    </div>

    <div class="row">
      <span class="row-label">{{ isImage ? "Border" : "Stroke" }}</span>
      <button
        type="button"
        class="pill"
        :aria-pressed="hasStroke"
        data-testid="style-stroke-toggle"
        @click="toggleStroke"
      >
        {{ hasStroke ? "On" : "Off" }}
      </button>
      <input
        type="color"
        :value="strokeColor"
        aria-label="Stroke colour"
        data-testid="style-stroke-color"
        @change="setStrokeColor"
      >
      <input
        type="number"
        class="width"
        :value="strokeWidth"
        :min="MIN_STROKE_WIDTH_PT"
        :max="MAX_STROKE_WIDTH_PT"
        step="0.25"
        aria-label="Stroke width (pt)"
        data-testid="style-stroke-width"
        @change="setStrokeWidth"
      >
    </div>
  </div>
</template>

<style scoped>
.style-popover {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-3);
  min-width: 14rem;
}

.row {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.row-label {
  width: 3.5rem;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.pill {
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-size: var(--text-sm);
}

.pill[aria-pressed="true"] {
  background: var(--accent-soft, var(--surface-sunken));
  border-color: var(--accent, var(--border-default));
}

.width {
  width: 4rem;
}

input[type="color"] {
  width: 2rem;
  height: 2rem;
  padding: 0;
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: transparent;
  cursor: pointer;
}
</style>
