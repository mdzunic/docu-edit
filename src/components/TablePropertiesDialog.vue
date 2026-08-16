<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

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
});

const emit = defineEmits(["apply", "cancel"]);

const borderStyle = ref(String(props.initial.borderStyle || "single"));
const borderWidthPt = ref(String(props.initial.borderWidthPt || 1));
const borderColor = ref(String(props.initial.borderColor || "#000000"));
const pendingPreset = ref("");

const shadingEnabled = ref(!!props.initial.shadingFill);
const shadingFill = ref(String(props.initial.shadingFill || "#ffff00"));

const paddingCm = ref(String(props.initial.paddingCm ?? 0.19));
const verticalAlign = ref(String(props.initial.verticalAlign || "top"));

const alignment = ref(String(props.initial.alignment || "left"));
const widthMode = ref(String(props.initial.widthMode || "fixed"));
const widthPercent = ref(String(props.initial.widthPercent || 100));

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
  const value = Number(paddingCm.value);
  return Number.isFinite(value) && value >= 0 && value <= 5 ? "" : "Padding must be between 0 and 5cm";
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
    shadingFill: shadingEnabled.value ? shadingFill.value : "",
    paddingCm: Number(paddingCm.value),
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
      >
    </section>

    <section class="section">
      <h3 class="section-title">
        Cell layout — {{ initial.scope }}
      </h3>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Padding (cm)</span>
          <input
            v-model="paddingCm"
            class="field-input"
            data-testid="table-props-padding"
            type="number"
            min="0"
            max="5"
            step="0.05"
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
      <button
        type="button"
        data-testid="table-props-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="table-props-apply"
        :disabled="!!error()"
        @click="apply"
      >
        OK
      </button>
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
  margin-bottom: 1.25rem;
  padding-bottom: 1rem;
  border-bottom: 1px solid #e0e0e0;
}

.section:last-of-type {
  border-bottom: none;
}

.section-title {
  margin: 0 0 0.5rem;
  font-size: 0.8125rem;
  font-weight: 600;
  color: #5f6368;
  text-transform: uppercase;
}

.field-row {
  display: flex;
  gap: 0.75rem;
  margin-bottom: 0.5rem;
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
}

.field-inline {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.5rem;
  font-size: 0.875rem;
}

.field-label {
  font-size: 0.8125rem;
  color: #3c4043;
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
}

.color-input {
  padding: 0.125rem;
  height: 2.2rem;
}

.preset-row {
  display: flex;
  gap: 0.5rem;
  flex-wrap: wrap;
}

.preset-button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.375rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.preset-button.selected {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
