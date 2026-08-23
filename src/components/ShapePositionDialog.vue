<script setup lang="ts">
import { ref, computed, watch } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import { parseLengthToCm, formatLength, asMeasurementUnit } from "../lib/core/units";

// The drawing "Position & size" dialog (SPEC-068 AC-42): the numeric-entry
// fallback for dragging a drawing handle, so a selected shape stays movable
// and resizable on a coarse pointer where a fine handle drag is impractical.
// Same dialog shape as ImageOptionsDialog.vue — a captured `initial`
// prefill, `apply`/`cancel` emits, local draft refs held in the display
// unit while the model keeps centimetres, OK disabled while invalid.
// Position (X/Y offset) only applies to a floating shape; for an inline
// shape the offset is meaningless (the wrap mode ignores it), so the
// position fields are disabled with a note.
const props = defineProps({
  /** Prefill: { offsetXCm, offsetYCm, widthCm, heightCm, inline }. */
  initial: { type: Object, required: true },
  /** The app-wide measurement unit (SPEC-064); fields display and parse in
   * it while the model keeps centimetres. */
  unit: { type: String, default: "cm" },
});

const emit = defineEmits(["apply", "cancel"]);

const measurementUnit = computed(() => asMeasurementUnit(props.unit));
const unitLabel = computed(() => props.unit);

const inline = props.initial.inline === true;

const offsetX = ref(formatLength(Number(props.initial.offsetXCm) || 0, measurementUnit.value));
const offsetY = ref(formatLength(Number(props.initial.offsetYCm) || 0, measurementUnit.value));
const width = ref(formatLength(Number(props.initial.widthCm), measurementUnit.value));
const height = ref(formatLength(Number(props.initial.heightCm), measurementUnit.value));

function widthCmValue() {
  return parseLengthToCm(width.value, measurementUnit.value);
}

function heightCmValue() {
  return parseLengthToCm(height.value, measurementUnit.value);
}

function offsetXCmValue() {
  return parseLengthToCm(offsetX.value, measurementUnit.value);
}

function offsetYCmValue() {
  return parseLengthToCm(offsetY.value, measurementUnit.value);
}

function widthError() {
  const value = widthCmValue();
  return value !== null && value > 0 && value <= 500
    ? ""
    : `Width must be between 0 and ${formatLength(500, measurementUnit.value)}${unitLabel.value}`;
}

function heightError() {
  const value = heightCmValue();
  // A Line legitimately has zero height, so the floor is 0 here, not >0.
  return value !== null && value >= 0 && value <= 500
    ? ""
    : `Height must be between 0 and ${formatLength(500, measurementUnit.value)}${unitLabel.value}`;
}

function positionError() {
  if (inline) {
    return "";
  }
  return offsetXCmValue() === null || offsetYCmValue() === null
    ? "Position must be a number"
    : "";
}

function error() {
  return widthError() || heightError() || positionError();
}

// Reformat the visible field text from the underlying centimetres when the
// app-wide unit changes while the dialog is open (SPEC-064).
watch(() => props.unit, (next, prev) => {
  for (const field of [offsetX, offsetY, width, height]) {
    const cm = parseLengthToCm(field.value, asMeasurementUnit(prev));
    if (cm !== null) {
      field.value = formatLength(cm, asMeasurementUnit(next));
    }
  }
});

function apply() {
  if (error()) {
    return;
  }
  emit("apply", {
    offsetXCm: inline ? 0 : (offsetXCmValue() ?? 0),
    offsetYCm: inline ? 0 : (offsetYCmValue() ?? 0),
    widthCm: widthCmValue(),
    heightCm: heightCmValue(),
  });
}
</script>

<template>
  <BaseDialog
    title="Position &amp; size"
    testid="shape-position-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <div class="field-row">
      <label class="field">
        <span class="field-label">Width ({{ unitLabel }})</span>
        <input
          v-model="width"
          class="field-input"
          data-testid="shape-position-width"
          type="text"
          inputmode="decimal"
          autofocus
        >
      </label>
      <label class="field">
        <span class="field-label">Height ({{ unitLabel }})</span>
        <input
          v-model="height"
          class="field-input"
          data-testid="shape-position-height"
          type="text"
          inputmode="decimal"
        >
      </label>
    </div>

    <div class="field-row">
      <label class="field">
        <span class="field-label">Horizontal ({{ unitLabel }})</span>
        <input
          v-model="offsetX"
          class="field-input"
          data-testid="shape-position-x"
          type="text"
          inputmode="decimal"
          :disabled="inline"
        >
      </label>
      <label class="field">
        <span class="field-label">Vertical ({{ unitLabel }})</span>
        <input
          v-model="offsetY"
          class="field-input"
          data-testid="shape-position-y"
          type="text"
          inputmode="decimal"
          :disabled="inline"
        >
      </label>
    </div>
    <p
      v-if="inline"
      class="field-note"
      data-testid="shape-position-inline-note"
    >
      An in-line drawing has no position offset; change the wrap mode to move it.
    </p>
    <p
      v-if="error()"
      class="dialog-error"
      data-testid="shape-position-error"
    >
      {{ error() }}
    </p>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="shape-position-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="shape-position-apply"
        :disabled="!!error()"
        @click="apply"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  margin-bottom: var(--space-4);
}

.field-row {
  display: flex;
  gap: var(--space-4);
}

.field-row .field {
  flex: 1;
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

.field-input:disabled {
  background: var(--surface-sunken);
  color: var(--text-disabled);
}

.field-note {
  margin: calc(-1 * var(--space-2)) 0 var(--space-4);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.dialog-error {
  margin: calc(-1 * var(--space-3)) 0 var(--space-4);
  color: var(--text-danger);
  font-size: var(--text-md);
}
</style>
