<script setup lang="ts">
import { ref, computed, watch } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import UiField from "./ui/UiField.vue";
import UiSelect from "./ui/UiSelect.vue";
import {
  parseLengthToCm,
  formatLength,
  asMeasurementUnit,
} from "../lib/core/units";
import { MIN_SHAPE_SIZE_CM } from "../lib/core/shapes";

// The drawing "Position & size" dialog (SPEC-072, expanding SPEC-068 AC-42):
// the numeric-entry path for every geometric property a drag adjusts — width,
// height, horizontal and vertical position, and rotation — so a selected
// drawing is fully editable from the keyboard and on a coarse pointer. Values
// display and parse in the app-wide measurement unit (SPEC-064) while the model
// keeps centimetres; the horizontal/vertical position each carries a
// "relative to" reference mapping to the model's relativeToH/relativeToV
// (SPEC-051). Size validates against MIN_SHAPE_SIZE_CM. Changes apply LIVE: each
// committed field emits `apply`, which the owner records as one undo entry.
// Position (offset) only applies to a floating shape.
const props = defineProps({
  /** Prefill: offsetXCm, offsetYCm, widthCm, heightCm, rotationDeg,
   * relativeToH, relativeToV, inline, isLine. */
  initial: { type: Object, required: true },
  /** The app-wide measurement unit (SPEC-064); fields display and parse in it
   * while the model keeps centimetres. */
  unit: { type: String, default: "cm" },
});

const emit = defineEmits(["apply", "dismiss"]);

const measurementUnit = computed(() => asMeasurementUnit(props.unit));
const unitLabel = computed(() => props.unit);

const inline = computed(() => props.initial.inline === true);
const isLine = computed(() => props.initial.isLine === true);

const width = ref("");
const height = ref("");
const offsetX = ref("");
const offsetY = ref("");
const rotation = ref("");
const relativeToH = ref("margin");
const relativeToV = ref("paragraph");
const lockAspect = ref(false);

// The horizontal reference maps to the model's relativeToH (page/margin/column,
// SPEC-051); vertical to relativeToV (page/margin/paragraph).
const H_REFERENCES = [
  { value: "margin", label: "Margin" },
  { value: "page", label: "Page" },
  { value: "column", label: "Column" },
];
const V_REFERENCES = [
  { value: "margin", label: "Margin" },
  { value: "page", label: "Page" },
  { value: "paragraph", label: "Paragraph" },
];

/** Re-seed every draft field from the (captured) model values. Runs on open and
 * whenever the selection changes under an open dialog (AC-51). */
function seed() {
  const unit = measurementUnit.value;
  width.value = formatLength(Number(props.initial.widthCm) || 0, unit);
  height.value = formatLength(Number(props.initial.heightCm) || 0, unit);
  offsetX.value = formatLength(Number(props.initial.offsetXCm) || 0, unit);
  offsetY.value = formatLength(Number(props.initial.offsetYCm) || 0, unit);
  rotation.value = String(Math.round(Number(props.initial.rotationDeg) || 0));
  relativeToH.value = String(props.initial.relativeToH ?? "margin");
  relativeToV.value = String(props.initial.relativeToV ?? "paragraph");
}
seed();
watch(() => props.initial, seed);

// Reformat the visible field text when the app-wide unit changes mid-dialog.
watch(() => props.unit, (next, prev) => {
  for (const field of [width, height, offsetX, offsetY]) {
    const cmValue = parseLengthToCm(field.value, asMeasurementUnit(prev));
    if (cmValue !== null) {
      field.value = formatLength(cmValue, asMeasurementUnit(next));
    }
  }
});

// Parsed cm value per field (null when unparseable). No-param computeds keep
// the .vue script free of the type annotations espree bars.
const widthCm = computed(() => parseLengthToCm(width.value, measurementUnit.value));
const heightCm = computed(() => parseLengthToCm(height.value, measurementUnit.value));
const offsetXCm = computed(() => parseLengthToCm(offsetX.value, measurementUnit.value));
const offsetYCm = computed(() => parseLengthToCm(offsetY.value, measurementUnit.value));

const widthError = computed(() => {
  const value = widthCm.value;
  return value !== null && value >= MIN_SHAPE_SIZE_CM
    ? ""
    : `Width must be at least ${formatLength(MIN_SHAPE_SIZE_CM, measurementUnit.value)}${unitLabel.value}`;
});

const heightError = computed(() => {
  const value = heightCm.value;
  // A Line legitimately has zero height (a horizontal rule); everything else is
  // floored at MIN_SHAPE_SIZE_CM (AC-44).
  const floor = isLine.value ? 0 : MIN_SHAPE_SIZE_CM;
  return value !== null && value >= floor
    ? ""
    : `Height must be at least ${formatLength(floor, measurementUnit.value)}${unitLabel.value}`;
});

const rotationError = computed(() => {
  const value = Number(rotation.value);
  return Number.isFinite(value) && value >= 0 && value <= 359
    ? ""
    : "Rotation must be 0–359";
});

/** The whole geometry in cm, or null when any field is invalid. */
function geometry() {
  const w = widthCm.value;
  const h = heightCm.value;
  const x = inline.value ? 0 : offsetXCm.value;
  const y = inline.value ? 0 : offsetYCm.value;
  const rotationDeg = Number(rotation.value);
  if (
    w === null || h === null || x === null || y === null ||
    widthError.value || heightError.value || rotationError.value
  ) {
    return null;
  }
  return {
    widthCm: w,
    heightCm: h,
    offsetXCm: x,
    offsetYCm: y,
    rotationDeg,
    relativeToH: relativeToH.value,
    relativeToV: relativeToV.value,
  };
}

/** Emit a live apply when the changed field leaves a valid geometry (AC-43).
 * One `apply` per committed field change => one undo entry (AC-45). */
function commit() {
  const geo = geometry();
  if (geo) {
    emit("apply", geo);
  }
}

// When aspect ratio is locked, editing one size field scales the other to keep
// the shape's proportions (AC-42), using the ratio captured at open.
function aspectRatio() {
  const w = Number(props.initial.widthCm) || 0;
  const h = Number(props.initial.heightCm) || 0;
  return w > 0 && h > 0 ? h / w : 0;
}

function commitWidth() {
  if (lockAspect.value && !widthError.value) {
    const w = widthCm.value;
    const ratio = aspectRatio();
    if (w !== null && ratio > 0) {
      height.value = formatLength(w * ratio, measurementUnit.value);
    }
  }
  commit();
}

function commitHeight() {
  if (lockAspect.value && !heightError.value) {
    const h = heightCm.value;
    const ratio = aspectRatio();
    if (h !== null && ratio > 0) {
      width.value = formatLength(h / ratio, measurementUnit.value);
    }
  }
  commit();
}
</script>

<template>
  <BaseDialog
    title="Position &amp; size"
    testid="shape-position-dialog"
    width="26rem"
    @dismiss="emit('dismiss')"
  >
    <fieldset class="group">
      <legend>Size</legend>
      <div class="field-row">
        <UiField
          :label="`Width (${unitLabel})`"
          for="shape-position-width"
          :error="widthError"
        >
          <input
            id="shape-position-width"
            v-model="width"
            class="field-input"
            data-testid="shape-position-width"
            type="text"
            inputmode="decimal"
            autofocus
            @change="commitWidth"
          >
        </UiField>
        <UiField
          :label="`Height (${unitLabel})`"
          for="shape-position-height"
          :error="heightError"
        >
          <input
            id="shape-position-height"
            v-model="height"
            class="field-input"
            data-testid="shape-position-height"
            type="text"
            inputmode="decimal"
            @change="commitHeight"
          >
        </UiField>
      </div>
      <label class="check">
        <input
          v-model="lockAspect"
          type="checkbox"
          data-testid="shape-position-lock-aspect"
        >
        Lock aspect ratio
      </label>
    </fieldset>

    <fieldset class="group">
      <legend>Position</legend>
      <div class="field-row">
        <UiField
          :label="`Horizontal (${unitLabel})`"
          for="shape-position-x"
        >
          <input
            id="shape-position-x"
            v-model="offsetX"
            class="field-input"
            data-testid="shape-position-x"
            type="text"
            inputmode="decimal"
            :disabled="inline"
            @change="commit"
          >
        </UiField>
        <UiField
          label="Relative to"
          for="shape-position-rel-h"
        >
          <UiSelect
            id="shape-position-rel-h"
            data-testid="shape-position-rel-h"
            :model-value="relativeToH"
            :options="H_REFERENCES"
            :disabled="inline"
            @update:model-value="(value) => { relativeToH = value; commit(); }"
          />
        </UiField>
      </div>
      <div class="field-row">
        <UiField
          :label="`Vertical (${unitLabel})`"
          for="shape-position-y"
        >
          <input
            id="shape-position-y"
            v-model="offsetY"
            class="field-input"
            data-testid="shape-position-y"
            type="text"
            inputmode="decimal"
            :disabled="inline"
            @change="commit"
          >
        </UiField>
        <UiField
          label="Relative to"
          for="shape-position-rel-v"
        >
          <UiSelect
            id="shape-position-rel-v"
            data-testid="shape-position-rel-v"
            :model-value="relativeToV"
            :options="V_REFERENCES"
            :disabled="inline"
            @update:model-value="(value) => { relativeToV = value; commit(); }"
          />
        </UiField>
      </div>
      <p
        v-if="inline"
        class="field-note"
        data-testid="shape-position-inline-note"
      >
        An in-line drawing has no position offset; change the wrap mode to move it.
      </p>
    </fieldset>

    <fieldset class="group">
      <legend>Rotation</legend>
      <UiField
        label="Degrees (0–359)"
        for="shape-position-rotation"
        :error="rotationError"
      >
        <input
          id="shape-position-rotation"
          v-model="rotation"
          class="field-input"
          data-testid="shape-position-rotation"
          type="text"
          inputmode="numeric"
          @change="commit"
        >
      </UiField>
    </fieldset>

    <template #actions>
      <UiButton
        variant="primary"
        data-testid="shape-position-done"
        @click="emit('dismiss')"
      >
        Done
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.group {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  padding: var(--space-3) var(--space-4) var(--space-1);
  margin: 0 0 var(--space-4);
}

.group legend {
  font-size: var(--text-sm);
  color: var(--text-secondary);
  padding: 0 var(--space-2);
}

.field-row {
  display: flex;
  gap: var(--space-4);
}

.field-row > * {
  flex: 1;
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

.field-input:disabled {
  background: var(--surface-sunken);
  color: var(--text-disabled);
}

.check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  margin: 0 0 var(--space-3);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-note {
  margin: 0 0 var(--space-3);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}
</style>
