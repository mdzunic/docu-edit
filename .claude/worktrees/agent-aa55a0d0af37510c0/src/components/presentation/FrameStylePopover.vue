<script setup lang="ts">
// The style popover for the selected frame (SPEC-082): fill colour (with "No
// fill") and stroke colour/width for shapes and text boxes; for images, stroke
// only. Each committed change is one undo entry — the parent records + persists
// on every `set-fill`/`set-stroke`. Runtime prop/emit declarations per the
// espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSlideFrame, asTheme } from "./slideRender";
import { MAX_STROKE_WIDTH_PT, MIN_STROKE_WIDTH_PT } from "../../lib/presentation/frames";
import {
  resolveColor,
  THEME_SWATCH_SLOTS,
  THEME_SWATCH_VARIANTS,
  themeSlotLabel,
} from "../../lib/core/theme";

const props = defineProps({
  /** The selected SlideFrame (shape / text / image). */
  frame: { type: Object, required: true },
  /** Optional DocumentTheme (SPEC-094): renders the theme swatch rows above
   * the colour inputs; a picked swatch stores the ThemeColorRef. Null keeps
   * SPEC-082's popover exactly as it was. */
  theme: { type: Object, default: null },
});

const emit = defineEmits(["set-fill", "set-stroke"]);

const frame = computed(() => asSlideFrame(props.frame));
const deckTheme = computed(() => asTheme(props.theme));

// SPEC-094: the theme swatch rows (base + five variants per slot), rendered
// only when a theme is passed. A swatch stores the REFERENCE: on a fillable
// frame it sets the fill; on an image it colours the border.
const themeRows = computed(() => {
  if (!props.theme) {
    return [];
  }
  const base = THEME_SWATCH_SLOTS.map((slot) => ({
    key: slot,
    name: themeSlotLabel(slot),
    hex: resolveColor({ theme: slot }, deckTheme.value),
    reference: { theme: slot },
  }));
  const variants = THEME_SWATCH_VARIANTS.map((variant) =>
    THEME_SWATCH_SLOTS.map((slot) => {
      const reference = Object({ theme: slot });
      if (variant.lumMod !== undefined) reference.lumMod = variant.lumMod;
      if (variant.lumOff !== undefined) reference.lumOff = variant.lumOff;
      return {
        key: `${slot}:${variant.label}`,
        name: `${themeSlotLabel(slot)}, ${variant.label}`,
        hex: resolveColor(reference, deckTheme.value),
        reference,
      };
    }),
  );
  return [base, ...variants];
});

function pickThemeSwatch(reference = Object()) {
  if (canFill.value) {
    emit("set-fill", { type: "solid", color: reference });
  } else {
    emit("set-stroke", { color: reference, widthPt: strokeWidth.value });
  }
}
const isImage = computed(() => frame.value.kind === "image");
const canFill = computed(() => frame.value.kind === "shape" || frame.value.kind === "text");

const fillColor = computed(() => {
  const f = frame.value;
  const color = (f.kind === "shape" || f.kind === "text") && f.fill?.type === "solid" ? f.fill.color : "#4472c4";
  // The native colour input needs a literal; a ref shows its resolved value.
  return resolveColor(color, deckTheme.value);
});
const hasFill = computed(() => {
  const f = frame.value;
  return (f.kind === "shape" || f.kind === "text") && f.fill?.type === "solid";
});

const stroke = computed(() => {
  const f = frame.value;
  return f.kind === "shape" || f.kind === "text" || f.kind === "image" ? f.stroke : undefined;
});
const strokeColor = computed(() => resolveColor(stroke.value?.color ?? "#000000", deckTheme.value));
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
      v-if="themeRows.length > 0"
      class="theme-block"
    >
      <span class="row-label">Theme</span>
      <div
        v-for="(row, rowIndex) in themeRows"
        :key="`theme-row-${rowIndex}`"
        class="theme-row"
      >
        <button
          v-for="cell in row"
          :key="cell.key"
          type="button"
          class="theme-swatch"
          data-testid="style-theme-swatch"
          :data-theme-slot="cell.reference.theme"
          :data-color="cell.hex"
          :title="cell.name"
          :aria-label="cell.name"
          :style="{ background: cell.hex }"
          @click="pickThemeSwatch(cell.reference)"
        />
      </div>
    </div>

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

.theme-block {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.theme-row {
  display: grid;
  grid-template-columns: repeat(10, 1rem);
  gap: var(--space-1);
}

.theme-swatch {
  width: 1rem;
  height: 1rem;
  padding: 0;
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  cursor: pointer;
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
