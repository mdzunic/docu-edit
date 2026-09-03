<script setup lang="ts">
import { computed } from "vue";
import UiSelect from "./ui/UiSelect.vue";
import {
  DEFAULT_NEW_SHAPE_OUTLINE,
  DEFAULT_SHAPE_SHADOW,
  SHAPE_SHADOW_PRESETS,
  SHAPE_STYLE_PRESETS,
  seededGradient,
  shapeFillBaseColor,
} from "../lib/core/shapes";
import { resolveColor } from "../lib/core/theme";
import { asDocumentTheme } from "../lib/documentModel";

// The full shape styling picker (SPEC-073 AC-16): style presets, the fill
// kind picker with its gradient editor and transparency slider, the widened
// outline controls (width/dash/join/cap, arrowheads for line-likes) and the
// shadow toggle + presets. Used from two hosts (AC-17): the contextual
// toolbar's Style popover and the Format → Drawing → Shape formatting…
// dialog — one component so the two cannot drift. Emits whole property
// values ("fill"/"outline"/"shadow" = one labelled undo entry each in the
// owner); a slider drag emits "preview-fill"/"preview-outline" live and the
// final value once on release (AC-25). Runtime prop declarations, not
// defineProps<{...}>() generics — the espree constraint every .vue script
// here shares.
const props = defineProps({
  /** Snapshot of the selection's styling: { fill, outline, shadow, isLine,
   * isConnector, isContainer } — fill/outline/shadow are the model values
   * (or undefined), the flags choose which sections show. */
  styling: { type: Object, required: true },
  /** The document's theme (SPEC-077) — a theme-referenced fill colour
   * resolves against it for the swatch and the gradient seed; applying
   * always re-emits the stored model value untouched. */
  theme: { type: Object, default: null },
});

const emit = defineEmits([
  "fill",
  "outline",
  "shadow",
  "preset",
  "preview-fill",
  "preview-outline",
  "pick-picture",
]);

const fill = computed(() => props.styling.fill);
const outline = computed(() => props.styling.outline);
const shadow = computed(() => props.styling.shadow);
const isLine = computed(() => props.styling.isLine === true);
const isConnector = computed(() => props.styling.isConnector === true);
const showFill = computed(() => !isLine.value && !isConnector.value);

const FILL_KINDS = [
  { value: "none", label: "No fill" },
  { value: "solid", label: "Solid" },
  { value: "gradient", label: "Gradient" },
  { value: "picture", label: "Picture" },
  { value: "pattern", label: "Pattern" },
];
const PATTERN_PRESETS = ["ltHorz", "ltVert", "ltUpDiag", "ltDnDiag", "smGrid", "diagCross"];
const DASHES = ["solid", "dash", "dot", "dashDot", "longDash", "longDashDot"];
const JOINS = ["miter", "round", "bevel"];
const CAPS = ["flat", "round", "square"];
const ARROW_KINDS = ["none", "triangle", "open", "stealth", "diamond", "oval"];
const ARROW_SIZES = ["small", "medium", "large"];

const fillKind = computed(() => (fill.value ? fill.value.kind : "none"));
const solidColor = computed(() => shapeFillBaseColor(fill.value) || "#4472c4");
// SPEC-077: the swatch input and the gradient seed need a concrete hex — a
// theme reference resolves for display only; applying keeps the model value.
const solidColorHex = computed(() => resolveColor(solidColor.value, asDocumentTheme(props.theme)));

/** The 0-100 transparency the slider shows (D5: transparency = 1 - alpha). */
function transparencyOf(alpha = 1) {
  return Math.round((1 - (typeof alpha === "number" ? alpha : 1)) * 100);
}
function alphaOf(transparency = 0) {
  const alpha = 1 - Math.min(100, Math.max(0, transparency)) / 100;
  return alpha >= 1 ? undefined : alpha;
}

const fillTransparency = computed(() => {
  const value = fill.value;
  if (!value) {
    return 0;
  }
  if (value.kind === "solid" || value.kind === "picture") {
    return transparencyOf(value.alpha);
  }
  return 0;
});

// --- fill edits -------------------------------------------------------------

function onFillKind(kind = "") {
  if (kind === fillKind.value) {
    return;
  }
  if (kind === "none") {
    emit("fill", undefined);
  } else if (kind === "solid") {
    emit("fill", { kind: "solid", color: solidColor.value });
  } else if (kind === "gradient") {
    // AC-22: seed from the current solid colour (or the accent) at 90°.
    emit("fill", { kind: "gradient", gradient: seededGradient(solidColorHex.value) });
  } else if (kind === "picture") {
    emit("pick-picture");
  } else if (kind === "pattern") {
    emit("fill", { kind: "pattern", preset: "ltDnDiag", foreground: solidColor.value, background: "#ffffff" });
  }
}

function patchedFillTransparency(event = new Event("input")) {
  const target = event.target;
  const value = fill.value;
  if (!(target instanceof HTMLInputElement) || !value) {
    return null;
  }
  if (value.kind !== "solid" && value.kind !== "picture") {
    return null;
  }
  const alpha = alphaOf(Number(target.value));
  const next = { ...value };
  if (alpha === undefined) {
    delete next.alpha;
  } else {
    next.alpha = alpha;
  }
  return next;
}

function onFillTransparencyInput(event = new Event("input")) {
  const next = patchedFillTransparency(event);
  if (next) {
    emit("preview-fill", next);
  }
}
function onFillTransparencyChange(event = new Event("change")) {
  const next = patchedFillTransparency(event);
  if (next) {
    emit("fill", next);
  }
}

function onSolidColor(event = new Event("input")) {
  const target = event.target;
  const value = fill.value;
  if (target instanceof HTMLInputElement) {
    const alpha = value && value.kind === "solid" ? value.alpha : undefined;
    emit(
      "fill",
      alpha === undefined ? { kind: "solid", color: target.value } : { kind: "solid", color: target.value, alpha },
    );
  }
}

// --- gradient editor (AC-16) ------------------------------------------------

const gradient = computed(() => (fill.value && fill.value.kind === "gradient" ? fill.value.gradient : null));

/** The gradient's stops as a real array (the styling prop arrives untyped),
 * so template iteration gets numeric indices and array semantics. */
function stopList() {
  const g = gradient.value;
  const list = g && Array.isArray(g.stops) ? [...g.stops] : [];
  return list;
}
const gradientStops = computed(() => stopList());

function emitGradient(next = {}) {
  emit("fill", { kind: "gradient", gradient: next });
}

function onGradientType(type = "") {
  const g = gradient.value;
  if (!g) {
    return;
  }
  const next = { ...g, type: type === "radial" ? "radial" : "linear" };
  if (next.type === "radial") {
    delete next.angleDeg;
  } else if (next.angleDeg === undefined) {
    next.angleDeg = 90;
  }
  emitGradient(next);
}

function onGradientAngle(event = new Event("change")) {
  const g = gradient.value;
  const target = event.target;
  if (!g || !(target instanceof HTMLInputElement)) {
    return;
  }
  const deg = Math.round(Number(target.value));
  if (Number.isFinite(deg)) {
    emitGradient({ ...g, angleDeg: ((deg % 360) + 360) % 360 });
  }
}

function replaceStop(index = 0, patch = {}) {
  const g = gradient.value;
  if (!g) {
    return;
  }
  const stops = stopList().map((stop, i) => (i === index ? { ...stop, ...patch } : stop));
  stops.sort((a, b) => Number(a.position) - Number(b.position));
  emitGradient({ ...g, stops });
}

function onStopColor(index = 0, event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    replaceStop(index, { color: target.value });
  }
}

function onStopPosition(index = 0, event = new Event("change")) {
  const g = gradient.value;
  const target = event.target;
  if (!g || !(target instanceof HTMLInputElement)) {
    return;
  }
  // First and last stops stay pinned at 0 and 1 (AC-3).
  if (index === 0 || index === g.stops.length - 1) {
    return;
  }
  const position = Math.min(1, Math.max(0, Number(target.value) / 100));
  if (Number.isFinite(position)) {
    replaceStop(index, { position });
  }
}

function onStopTransparency(index = 0, event = new Event("change")) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }
  const g = gradient.value;
  if (!g) {
    return;
  }
  const alpha = alphaOf(Number(target.value));
  const stop = { ...g.stops[index] };
  if (alpha === undefined) {
    delete stop.alpha;
  } else {
    stop.alpha = alpha;
  }
  const stops = stopList().map((entry, i) => (i === index ? stop : entry));
  emitGradient({ ...g, stops });
}

function addStop() {
  const g = gradient.value;
  if (!g) {
    return;
  }
  const stops = [...stopList()];
  const mid = { position: 0.5, color: stops[0].color };
  stops.push(mid);
  stops.sort((a, b) => Number(a.position) - Number(b.position));
  emitGradient({ ...g, stops });
}

function removeStop(index = 0) {
  const g = gradient.value;
  if (!g || stopList().length <= 2) {
    return;
  }
  const stops = stopList().filter((unused, i) => i !== index);
  stops[0] = { ...stops[0], position: 0 };
  stops[stops.length - 1] = { ...stops[stops.length - 1], position: 1 };
  emitGradient({ ...g, stops });
}

// --- picture / pattern ------------------------------------------------------

function onPictureMode(mode = "") {
  const value = fill.value;
  if (value && value.kind === "picture") {
    emit("fill", { ...value, mode: mode === "tile" ? "tile" : "stretch" });
  }
}

function onPatternPart(part = "preset", event = new Event("input")) {
  const value = fill.value;
  const target = event.target;
  if (!value || value.kind !== "pattern") {
    return;
  }
  if (part === "preset" && typeof event === "string") {
    return;
  }
  if (target instanceof HTMLInputElement || target instanceof HTMLSelectElement) {
    emit("fill", { ...value, [part]: target.value });
  }
}

function onPatternPreset(preset = "") {
  const value = fill.value;
  if (value && value.kind === "pattern") {
    emit("fill", { ...value, preset });
  }
}

// --- outline ----------------------------------------------------------------

function baseOutline() {
  return outline.value ? { ...outline.value } : { ...DEFAULT_NEW_SHAPE_OUTLINE };
}

function emitOutline(patch = {}) {
  emit("outline", { ...baseOutline(), ...patch });
}

function onOutlineColor(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emitOutline({ color: target.value });
  }
}

function onOutlineWidth(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    const widthPt = Number(target.value);
    if (Number.isFinite(widthPt) && widthPt > 0) {
      emitOutline({ widthPt });
    }
  }
}

function onOutlineDash(dash = "solid") {
  emitOutline({ dash });
}
function onOutlineJoin(join = "miter") {
  const next = baseOutline();
  if (join === "miter") {
    delete next.join;
  } else {
    next.join = join;
  }
  emit("outline", next);
}
function onOutlineCap(cap = "flat") {
  const next = baseOutline();
  if (cap === "flat") {
    delete next.cap;
  } else {
    next.cap = cap;
  }
  emit("outline", next);
}

const outlineTransparency = computed(() => transparencyOf(outline.value ? outline.value.alpha : 1));

function patchedOutlineTransparency(event = new Event("input")) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || !outline.value) {
    return null;
  }
  const alpha = alphaOf(Number(target.value));
  const next = baseOutline();
  if (alpha === undefined) {
    delete next.alpha;
  } else {
    next.alpha = alpha;
  }
  return next;
}

function onOutlineTransparencyInput(event = new Event("input")) {
  const next = patchedOutlineTransparency(event);
  if (next) {
    emit("preview-outline", next);
  }
}
function onOutlineTransparencyChange(event = new Event("change")) {
  const next = patchedOutlineTransparency(event);
  if (next) {
    emit("outline", next);
  }
}

// --- arrowheads (AC-14) -----------------------------------------------------

const showArrows = computed(() => isLine.value || isConnector.value);

function arrowOf(end = "start") {
  const current = outline.value;
  const arrow = current ? (end === "start" ? current.startArrow : current.endArrow) : undefined;
  return arrow || { kind: "none", size: "medium" };
}

function onArrowKind(end = "start", kind = "none") {
  const next = baseOutline();
  const key = end === "start" ? "startArrow" : "endArrow";
  if (kind === "none") {
    delete next[key];
  } else {
    // Turning an arrowhead on with no size chosen uses medium (AC-14).
    next[key] = { kind, size: arrowOf(end).size || "medium" };
  }
  emit("outline", next);
}

function onArrowSize(end = "start", size = "medium") {
  const next = baseOutline();
  const key = end === "start" ? "startArrow" : "endArrow";
  const current = arrowOf(end);
  if (current.kind === "none") {
    return;
  }
  next[key] = { kind: current.kind, size };
  emit("outline", next);
}

// --- shadow (AC-20/AC-21) ---------------------------------------------------

const shadowPresetId = computed(() => {
  const current = shadow.value;
  if (!current) {
    return "none";
  }
  const match = SHAPE_SHADOW_PRESETS.find(
    (preset) =>
      preset.shadow &&
      preset.shadow.angleDeg === current.angleDeg &&
      preset.shadow.color === current.color &&
      preset.shadow.alpha === current.alpha &&
      preset.shadow.blurPt === current.blurPt &&
      preset.shadow.distancePt === current.distancePt,
  );
  return match ? match.id : "custom";
});

function onShadowToggle() {
  emit("shadow", shadow.value ? undefined : { ...DEFAULT_SHAPE_SHADOW });
}

function onShadowPreset(id = "none") {
  const preset = SHAPE_SHADOW_PRESETS.find((entry) => entry.id === id);
  emit("shadow", preset && preset.shadow ? { ...preset.shadow } : undefined);
}
</script>

<template>
  <div
    class="shape-style-panel"
    data-testid="shape-style-panel"
  >
    <section
      v-if="!isConnector"
      class="panel-section"
      aria-label="Style presets"
    >
      <span class="section-label">Style</span>
      <div
        class="preset-row"
        data-testid="shape-style-presets"
      >
        <button
          v-for="preset in SHAPE_STYLE_PRESETS"
          :key="preset.id"
          type="button"
          class="preset-button"
          :data-testid="`shape-style-preset-${preset.id}`"
          :title="preset.label"
          @click="emit('preset', preset.id)"
        >
          {{ preset.label }}
        </button>
      </div>
    </section>

    <section
      v-if="showFill"
      class="panel-section"
      aria-label="Fill"
    >
      <span class="section-label">Fill</span>
      <UiSelect
        :model-value="fillKind"
        :options="FILL_KINDS"
        data-testid="shape-fill-kind"
        label="Fill kind"
        @update:model-value="onFillKind"
      />
      <label
        v-if="fillKind === 'solid'"
        class="inline-field"
      >
        Colour
        <input
          type="color"
          data-testid="shape-fill-solid-color"
          :value="solidColorHex"
          aria-label="Fill colour"
          @input="onSolidColor"
        >
      </label>

      <div
        v-if="gradient"
        class="gradient-editor"
        data-testid="shape-gradient-editor"
      >
        <UiSelect
          :model-value="gradient.type"
          :options="[{ value: 'linear', label: 'Linear' }, { value: 'radial', label: 'Radial' }]"
          data-testid="shape-gradient-type"
          label="Gradient type"
          @update:model-value="onGradientType"
        />
        <label
          v-if="gradient.type === 'linear'"
          class="inline-field"
        >
          Angle
          <input
            type="number"
            class="num"
            data-testid="shape-gradient-angle"
            :value="gradient.angleDeg ?? 0"
            min="0"
            max="359"
            aria-label="Gradient angle in degrees"
            @change="onGradientAngle"
          >
        </label>
        <div
          v-for="(stop, index) in gradientStops"
          :key="index"
          class="gradient-stop"
          :data-testid="`shape-gradient-stop-${index}`"
        >
          <input
            type="color"
            :value="stop.color"
            :data-testid="`shape-gradient-stop-color-${index}`"
            :aria-label="`Stop ${index + 1} colour`"
            @input="onStopColor(index, $event)"
          >
          <input
            type="number"
            class="num"
            :value="Math.round(stop.position * 100)"
            min="0"
            max="100"
            :disabled="index === 0 || index === gradientStops.length - 1"
            :aria-label="`Stop ${index + 1} position %`"
            @change="onStopPosition(index, $event)"
          >
          <input
            type="number"
            class="num"
            :value="Math.round((1 - (stop.alpha ?? 1)) * 100)"
            min="0"
            max="100"
            :aria-label="`Stop ${index + 1} transparency %`"
            @change="onStopTransparency(index, $event)"
          >
          <button
            type="button"
            class="stop-remove"
            :disabled="gradientStops.length <= 2"
            :data-testid="`shape-gradient-stop-remove-${index}`"
            :aria-label="`Remove stop ${index + 1}`"
            @click="removeStop(index)"
          >
            ×
          </button>
        </div>
        <button
          type="button"
          data-testid="shape-gradient-add-stop"
          @click="addStop"
        >
          Add stop
        </button>
      </div>

      <template v-if="fill && fill.kind === 'picture'">
        <UiSelect
          :model-value="fill.mode"
          :options="[{ value: 'stretch', label: 'Stretch' }, { value: 'tile', label: 'Tile' }]"
          data-testid="shape-picture-mode"
          label="Picture fill mode"
          @update:model-value="onPictureMode"
        />
        <button
          type="button"
          data-testid="shape-picture-choose"
          @click="emit('pick-picture')"
        >
          Choose picture…
        </button>
      </template>

      <template v-if="fill && fill.kind === 'pattern'">
        <UiSelect
          :model-value="fill.preset"
          :options="PATTERN_PRESETS.map((value) => ({ value, label: value }))"
          data-testid="shape-pattern-preset"
          aria-label="Pattern preset"
          @update:model-value="onPatternPreset"
        />
        <label class="inline-field">
          Fore
          <input
            type="color"
            data-testid="shape-pattern-foreground"
            :value="fill.foreground"
            aria-label="Pattern foreground colour"
            @input="onPatternPart('foreground', $event)"
          >
        </label>
        <label class="inline-field">
          Back
          <input
            type="color"
            data-testid="shape-pattern-background"
            :value="fill.background"
            aria-label="Pattern background colour"
            @input="onPatternPart('background', $event)"
          >
        </label>
      </template>

      <label
        v-if="fill && (fill.kind === 'solid' || fill.kind === 'picture')"
        class="inline-field slider-field"
      >
        Transparency
        <input
          type="range"
          min="0"
          max="100"
          :value="fillTransparency"
          data-testid="shape-fill-transparency"
          aria-label="Fill transparency"
          @input="onFillTransparencyInput"
          @change="onFillTransparencyChange"
        >
      </label>
    </section>

    <section
      class="panel-section"
      aria-label="Outline"
    >
      <span class="section-label">Outline</span>
      <label class="inline-field">
        Colour
        <input
          type="color"
          data-testid="shape-style-outline-color"
          :value="outline ? outline.color : '#2e5395'"
          aria-label="Outline colour"
          @input="onOutlineColor"
        >
      </label>
      <label class="inline-field">
        Width
        <input
          type="number"
          class="num"
          min="0.25"
          step="0.25"
          data-testid="shape-style-outline-width"
          :value="outline ? outline.widthPt : 1"
          aria-label="Outline width in points"
          @change="onOutlineWidth"
        >
      </label>
      <UiSelect
        :model-value="outline ? outline.dash : 'solid'"
        :options="DASHES.map((value) => ({ value, label: value }))"
        data-testid="shape-style-outline-dash"
        aria-label="Outline dash"
        @update:model-value="onOutlineDash"
      />
      <UiSelect
        :model-value="outline && outline.join ? outline.join : 'miter'"
        :options="JOINS.map((value) => ({ value, label: value }))"
        data-testid="shape-outline-join"
        aria-label="Outline join"
        @update:model-value="onOutlineJoin"
      />
      <UiSelect
        :model-value="outline && outline.cap ? outline.cap : 'flat'"
        :options="CAPS.map((value) => ({ value, label: value }))"
        data-testid="shape-outline-cap"
        aria-label="Outline cap"
        @update:model-value="onOutlineCap"
      />
      <label class="inline-field slider-field">
        Transparency
        <input
          type="range"
          min="0"
          max="100"
          :value="outlineTransparency"
          :disabled="!outline"
          data-testid="shape-outline-transparency"
          aria-label="Outline transparency"
          @input="onOutlineTransparencyInput"
          @change="onOutlineTransparencyChange"
        >
      </label>
    </section>

    <section
      v-if="showArrows"
      class="panel-section"
      aria-label="Arrowheads"
    >
      <span class="section-label">Arrows</span>
      <div class="arrow-row">
        <UiSelect
          :model-value="arrowOf('start').kind"
          :options="ARROW_KINDS.map((value) => ({ value, label: `Start: ${value}` }))"
          data-testid="shape-start-arrow-kind"
          aria-label="Start arrowhead"
          @update:model-value="(kind) => onArrowKind('start', kind)"
        />
        <UiSelect
          :model-value="arrowOf('start').size"
          :options="ARROW_SIZES.map((value) => ({ value, label: value }))"
          data-testid="shape-start-arrow-size"
          aria-label="Start arrowhead size"
          @update:model-value="(size) => onArrowSize('start', size)"
        />
      </div>
      <div class="arrow-row">
        <UiSelect
          :model-value="arrowOf('end').kind"
          :options="ARROW_KINDS.map((value) => ({ value, label: `End: ${value}` }))"
          data-testid="shape-end-arrow-kind"
          aria-label="End arrowhead"
          @update:model-value="(kind) => onArrowKind('end', kind)"
        />
        <UiSelect
          :model-value="arrowOf('end').size"
          :options="ARROW_SIZES.map((value) => ({ value, label: value }))"
          data-testid="shape-end-arrow-size"
          aria-label="End arrowhead size"
          @update:model-value="(size) => onArrowSize('end', size)"
        />
      </div>
    </section>

    <section
      v-if="!isConnector"
      class="panel-section"
      aria-label="Shadow"
    >
      <span class="section-label">Shadow</span>
      <button
        type="button"
        data-testid="shape-shadow-toggle"
        :aria-pressed="!!shadow"
        @click="onShadowToggle"
      >
        {{ shadow ? "Remove shadow" : "Add shadow" }}
      </button>
      <UiSelect
        :model-value="shadowPresetId"
        :options="SHAPE_SHADOW_PRESETS.map((preset) => ({ value: preset.id, label: preset.label }))"
        data-testid="shape-shadow-preset"
        aria-label="Shadow preset"
        @update:model-value="onShadowPreset"
      />
    </section>
  </div>
</template>

<style scoped>
.shape-style-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
  min-width: 260px;
  /* The popover host sits mid-viewport; a hard cap keeps every control
     reachable (the panel scrolls) rather than spilling past the viewport
     bottom where a click cannot land. */
  max-height: min(46vh, 340px);
  overflow-y: auto;
  padding: var(--space-2) var(--space-1);
}
.panel-section {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  border-top: var(--border-width) solid var(--border-subtle);
  padding-top: var(--space-2);
}
.panel-section:first-child {
  border-top: none;
  padding-top: 0;
}
.section-label {
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  text-transform: uppercase;
  letter-spacing: 0.04em;
}
.preset-row {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}
.preset-button {
  font: inherit;
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-lg);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}
.preset-button:hover {
  background: var(--surface-selected);
}
.inline-field {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-sm);
}
.inline-field input[type="color"] {
  inline-size: var(--space-7);
  block-size: var(--space-6);
  padding: 0;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-lg);
  background: none;
}
.num {
  inline-size: 56px;
  font: inherit;
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-2);
}
.slider-field input[type="range"] {
  flex: 1;
}
.gradient-editor {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}
.gradient-stop {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}
.stop-remove {
  font: inherit;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-lg);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
  padding: 0 var(--space-2);
}
.arrow-row {
  display: flex;
  gap: var(--space-2);
}
</style>
