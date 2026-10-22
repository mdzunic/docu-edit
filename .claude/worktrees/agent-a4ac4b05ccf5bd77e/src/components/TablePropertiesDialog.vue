<script setup lang="ts">
import { computed, ref, watch } from "vue";
import { asMeasurementUnit, formatLength, parseLengthToCm } from "../lib/core/units";
import { isThemeColor, resolveColor } from "../lib/core/theme";
import { asDocumentTheme, plainModelColor } from "../lib/documentModel";
import BaseDialog from "./BaseDialog.vue";
import ThemeSwatchRow from "./ThemeSwatchRow.vue";
import UiButton from "./ui/UiButton.vue";

// The table properties dialog (SPEC-023 AC-10..23): border style/width/
// colour with the four standard presets, cell shading, padding, vertical
// alignment, table alignment and width. Every field commits together on OK
// as one combined edit (AC-34 — one undo step); Cancel discards everything.
// Border presets (AC-18..21) are buttons that stage which preset OK will
// apply — style/width/colour are read fresh at OK time, so changing them
// after clicking a preset still uses the latest values. BaseDialog owns the
// shell, Escape and initial focus (the [autofocus] border-style select).
const props = defineProps({
  /** Prefill: { borderStyle, borderWidthPt, borderColor, shadingFill,
   * paddingCm, verticalAlign, alignment, widthMode, widthPercent, scope }.
   * `shadingFill` may be "" for no shading; `scope` is a display string
   * ("the whole table" / "2 selected cells") shown next to the border/
   * shading/padding/alignment fields (AC-13, AC-14). */
  initial: { type: Object, required: true },
  /** The document's measurement-unit preference (SPEC-064) — the length
   * fields display and parse in this unit; the model stays in centimetres. */
  unit: { type: String, default: "cm" },
  /** The DocumentTheme (SPEC-077 AC-22): when present, a theme swatch row
   * renders under the shading colour input; picking a swatch makes OK carry
   * the ThemeColorRef itself, not the hex it displays as. Null (the default)
   * renders no row — existing mounts are unchanged. */
  theme: { type: Object, default: null },
});

const emit = defineEmits(["apply", "cancel"]);

// The unit the length fields display/parse in; the stored padding stays cm.
const measurementUnit = computed(() => asMeasurementUnit(props.unit));
const unitLabel = computed(() => props.unit);

const borderStyle = ref(String(props.initial.borderStyle || "single"));
const borderWidthPt = ref(String(props.initial.borderWidthPt || 1));
const borderColor = ref(String(props.initial.borderColor || "#000000"));
const pendingPreset = ref("");

const shadingEnabled = ref(!!props.initial.shadingFill);
// SPEC-077 AC-22/23: a theme swatch stores the reference here so OK carries
// it to the cells; any manual edit of the colour input clears it back to a
// literal. The colour input itself only ever displays a resolved hex.
const shadingRef = ref(isThemeColor(props.initial.shadingFill) ? props.initial.shadingFill : null);
const shadingFill = ref(
  isThemeColor(props.initial.shadingFill)
    ? resolveColor(props.initial.shadingFill, asDocumentTheme(props.theme))
    : String(props.initial.shadingFill || "#ffff00"),
);

/** A theme swatch (SPEC-077): store the reference, show its resolved hex. */
function onShadingSwatch(reference = Object()) {
  shadingRef.value = reference;
  shadingEnabled.value = true;
  shadingFill.value = resolveColor(reference, asDocumentTheme(props.theme));
}

// Held as display text in the current unit; the underlying cm stays canonical.
const paddingText = ref(formatLength(props.initial.paddingCm ?? 0.19, measurementUnit.value));
const verticalAlign = ref(String(props.initial.verticalAlign || "top"));

const alignment = ref(String(props.initial.alignment || "left"));
const widthMode = ref(String(props.initial.widthMode || "fixed"));
const widthPercent = ref(String(props.initial.widthPercent || 100));

// Reformat the visible padding text when the unit changes, re-deriving it from
// its underlying cm value so centimetres stay the source of truth (SPEC-064).
watch(
  () => props.unit,
  (next = "cm", previous = "cm") => {
    const paddingCm = parseLengthToCm(paddingText.value, asMeasurementUnit(previous));
    if (paddingCm !== null) {
      paddingText.value = formatLength(paddingCm, asMeasurementUnit(next));
    }
  },
);

const BORDER_PRESETS = [
  { key: "all", label: "All borders" },
  { key: "outer", label: "Outer only" },
  { key: "inner", label: "Inner only" },
  { key: "none", label: "None" },
];

function choosePreset(preset = "") {
  pendingPreset.value = preset;
}

function widthError() {
  const value = Number(borderWidthPt.value);
  return Number.isFinite(value) && value > 0 && value <= 96 ? "" : "Border width must be between 0 and 96pt";
}

function paddingError() {
  const value = parseLengthToCm(paddingText.value, measurementUnit.value);
  return value !== null && value >= 0 && value <= 5
    ? ""
    : `Padding must be between ${formatLength(0, measurementUnit.value)} and ${formatLength(5, measurementUnit.value)}${unitLabel.value}`;
}

function percentError() {
  if (widthMode.value !== "percent") {
    return "";
  }
  const value = Number(widthPercent.value);
  return Number.isFinite(value) && value >= 1 && value <= 100 ? "" : "Width must be between 1 and 100%";
}

function error() {
  return widthError() || paddingError() || percentError();
}

function apply() {
  if (error()) {
    return;
  }
  emit("apply", {
    borderPreset: pendingPreset.value,
    borderStyle: borderStyle.value,
    borderWidthPt: Number(borderWidthPt.value),
    borderColor: borderColor.value,
    // A stored theme reference wins over the hex it displays as (SPEC-077);
    // plainModelColor detaches it from this ref's reactive wrapper so the
    // model never holds a Proxy (see documentModel.ts).
    shadingFill: shadingEnabled.value ? (plainModelColor(shadingRef.value) ?? shadingFill.value) : "",
    paddingCm: parseLengthToCm(paddingText.value, measurementUnit.value) ?? 0,
    verticalAlign: verticalAlign.value,
    alignment: alignment.value,
    widthMode: widthMode.value,
    widthPercent: Number(widthPercent.value),
  });
}
</script>

<template>
  <BaseDialog
    title="Table properties"
    testid="table-properties-dialog"
    width="28rem"
    @dismiss="emit('cancel')"
  >
    <section class="section">
      <h3 class="section-title">
        Borders — {{ initial.scope }}
      </h3>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Style</span>
          <select
            v-model="borderStyle"
            class="field-input"
            data-testid="table-props-border-style"
            autofocus
          >
            <option value="single">
              Single
            </option>
            <option value="double">
              Double
            </option>
            <option value="dashed">
              Dashed
            </option>
            <option value="dotted">
              Dotted
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Width (pt)</span>
          <input
            v-model="borderWidthPt"
            class="field-input"
            data-testid="table-props-border-width"
            type="number"
            min="0.25"
            max="96"
            step="0.25"
          >
        </label>
        <label class="field">
          <span class="field-label">Colour</span>
          <input
            v-model="borderColor"
            class="field-input color-input"
            data-testid="table-props-border-color"
            type="color"
          >
        </label>
      </div>
      <div
        class="preset-row"
        role="group"
        aria-label="Border presets"
      >
        <button
          v-for="preset in BORDER_PRESETS"
          :key="preset.key"
          type="button"
          class="preset-button"
          :class="{ selected: pendingPreset === preset.key }"
          :data-testid="`table-props-preset-${preset.key}`"
          :aria-pressed="pendingPreset === preset.key"
          @click="choosePreset(preset.key)"
        >
          {{ preset.label }}
        </button>
      </div>
      <p
        v-if="widthError()"
        class="dialog-error"
        data-testid="table-props-width-error"
      >
        {{ widthError() }}
      </p>
    </section>

    <section class="section">
      <h3 class="section-title">
        Shading — {{ initial.scope }}
      </h3>
      <label class="field-inline">
        <input
          v-model="shadingEnabled"
          type="checkbox"
          data-testid="table-props-shading-enabled"
        >
        <span>Background colour</span>
      </label>
      <input
        v-model="shadingFill"
        class="field-input color-input"
        data-testid="table-props-shading-color"
        type="color"
        :disabled="!shadingEnabled"
        @input="shadingRef = null"
      >
      <ThemeSwatchRow
        :theme="theme"
        testid="table-props-shading-theme"
        @select="onShadingSwatch"
      />
    </section>

    <section class="section">
      <h3 class="section-title">
        Cell layout — {{ initial.scope }}
      </h3>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Padding ({{ unitLabel }})</span>
          <input
            v-model="paddingText"
            class="field-input"
            data-testid="table-props-padding"
            type="text"
            inputmode="decimal"
          >
        </label>
        <label class="field">
          <span class="field-label">Vertical alignment</span>
          <select
            v-model="verticalAlign"
            class="field-input"
            data-testid="table-props-valign"
          >
            <option value="top">
              Top
            </option>
            <option value="center">
              Middle
            </option>
            <option value="bottom">
              Bottom
            </option>
          </select>
        </label>
      </div>
      <p
        v-if="paddingError()"
        class="dialog-error"
        data-testid="table-props-padding-error"
      >
        {{ paddingError() }}
      </p>
    </section>

    <section class="section">
      <h3 class="section-title">
        Table
      </h3>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Alignment</span>
          <select
            v-model="alignment"
            class="field-input"
            data-testid="table-props-alignment"
          >
            <option value="left">
              Left
            </option>
            <option value="center">
              Centre
            </option>
            <option value="right">
              Right
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Width</span>
          <select
            v-model="widthMode"
            class="field-input"
            data-testid="table-props-width-mode"
          >
            <option value="auto">
              Auto
            </option>
            <option value="fixed">
              Fixed
            </option>
            <option value="percent">
              Percentage of text area
            </option>
          </select>
        </label>
        <label
          v-if="widthMode === 'percent'"
          class="field"
        >
          <span class="field-label">Percent</span>
          <input
            v-model="widthPercent"
            class="field-input"
            data-testid="table-props-width-percent"
            type="number"
            min="1"
            max="100"
          >
        </label>
      </div>
      <p
        v-if="percentError()"
        class="dialog-error"
        data-testid="table-props-percent-error"
      >
        {{ percentError() }}
      </p>
    </section>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="table-props-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="table-props-apply"
        :disabled="!!error()"
        @click="apply"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* Four sections make this dialog tall, so the panel keeps its own scroll
   cap (the base's .dialog is the short-dialog variant without one). */
:deep(.dialog) {
  max-height: 85vh;
  overflow-y: auto;
}

.section {
  margin-bottom: var(--space-6);
  padding-bottom: var(--space-5);
  border-bottom: var(--border-width) solid var(--border-default);
}

.section:last-of-type {
  border-bottom: none;
}

.section-title {
  margin: 0 0 var(--space-3);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  text-transform: uppercase;
}

.field-row {
  display: flex;
  gap: var(--space-4);
  margin-bottom: var(--space-3);
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.field-inline {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-bottom: var(--space-3);
  font-size: var(--text-md);
}

.field-label {
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.color-input {
  padding: var(--space-1);
  height: 2.2rem;
}

.preset-row {
  display: flex;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.preset-button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-3) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.preset-button.selected {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-md);
}

</style>
