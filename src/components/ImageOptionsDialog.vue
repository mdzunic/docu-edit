<script setup lang="ts">
import { ref, computed, watch } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import { parseLengthToCm, formatLength, asMeasurementUnit } from "../lib/core/units";

// The "Image options" panel (SPEC-024 AC-15..22): alt text, size in cm with
// a lock-aspect-ratio toggle, and a "Reset size" button. Same dialog shape
// as TablePropertiesDialog.vue — a captured `initial` prefill, `apply`/
// `cancel` emits, local draft refs, OK disabled while invalid. BaseDialog
// owns the shell, Escape and initial focus (the [autofocus] alt-text
// input).
const props = defineProps({
  /** Prefill: { altText, widthCm, heightCm, naturalWidthCm, naturalHeightCm,
   * lockAspectRatio }. */
  initial: { type: Object, required: true },
  /** The app-wide measurement unit (SPEC-064); size fields display and parse
   * in it while the model keeps centimetres. */
  unit: { type: String, default: "cm" },
});

const emit = defineEmits(["apply", "cancel"]);

// The size fields show/enter values in the current unit; unitLabel names it.
const measurementUnit = computed(() => asMeasurementUnit(props.unit));
const unitLabel = computed(() => props.unit);

const altText = ref(String(props.initial.altText || ""));
// widthCm/heightCm hold the field text in the display unit; widthCmValue()/
// heightCmValue() convert back to the centimetres the model stores.
const widthCm = ref(formatLength(Number(props.initial.widthCm), measurementUnit.value));
const heightCm = ref(formatLength(Number(props.initial.heightCm), measurementUnit.value));
const lockAspectRatio = ref(props.initial.lockAspectRatio !== false);

const naturalWidthCm = Number(props.initial.naturalWidthCm) || Number(props.initial.widthCm) || 1;
const naturalHeightCm = Number(props.initial.naturalHeightCm) || Number(props.initial.heightCm) || 1;
// The ratio lock preserves — SPEC-024 AC-19's "original aspect ratio" is the
// image's natural (source) ratio, not whatever the box currently shows.
const naturalRatio = naturalWidthCm / naturalHeightCm;

function widthCmValue() {
  return parseLengthToCm(widthCm.value, measurementUnit.value);
}

function heightCmValue() {
  return parseLengthToCm(heightCm.value, measurementUnit.value);
}

function widthError() {
  const value = widthCmValue();
  return value !== null && value > 0 && value <= 500
    ? ""
    : `Width must be between 0 and ${formatLength(500, measurementUnit.value)}${unitLabel.value}`;
}

function heightError() {
  const value = heightCmValue();
  return value !== null && value > 0 && value <= 500
    ? ""
    : `Height must be between 0 and ${formatLength(500, measurementUnit.value)}${unitLabel.value}`;
}

function error() {
  return widthError() || heightError();
}

// AC-19: locked, changing one dimension updates the other to preserve the
// original aspect ratio. AC-20: unlocked, they are independent.
function handleWidthInput() {
  if (!lockAspectRatio.value) {
    return;
  }
  const width = widthCmValue();
  if (width !== null && width > 0) {
    heightCm.value = formatLength(width / naturalRatio, measurementUnit.value);
  }
}

function handleHeightInput() {
  if (!lockAspectRatio.value) {
    return;
  }
  const height = heightCmValue();
  if (height !== null && height > 0) {
    widthCm.value = formatLength(height * naturalRatio, measurementUnit.value);
  }
}

// AC-21, AC-22: restores natural size and re-enables the lock.
function resetSize() {
  widthCm.value = formatLength(naturalWidthCm, measurementUnit.value);
  heightCm.value = formatLength(naturalHeightCm, measurementUnit.value);
  lockAspectRatio.value = true;
}

// Reformat the visible field text from the underlying centimetres when the
// app-wide unit changes while the dialog is open (SPEC-064).
watch(() => props.unit, (next, prev) => {
  const width = parseLengthToCm(widthCm.value, asMeasurementUnit(prev));
  const height = parseLengthToCm(heightCm.value, asMeasurementUnit(prev));
  if (width !== null) {
    widthCm.value = formatLength(width, asMeasurementUnit(next));
  }
  if (height !== null) {
    heightCm.value = formatLength(height, asMeasurementUnit(next));
  }
});

function apply() {
  if (error()) {
    return;
  }
  emit("apply", {
    altText: altText.value.trim(),
    widthCm: widthCmValue(),
    heightCm: heightCmValue(),
  });
}
</script>

<template>
  <BaseDialog
    title="Image options"
    testid="image-options-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span class="field-label">Alt text (description)</span>
      <input
        v-model="altText"
        class="field-input"
        data-testid="image-options-alt-text"
        type="text"
        placeholder="Describe this image"
        autofocus
      >
    </label>

    <div class="field-row">
      <label class="field">
        <span class="field-label">Width ({{ unitLabel }})</span>
        <input
          v-model="widthCm"
          class="field-input"
          data-testid="image-options-width"
          type="number"
          min="0.1"
          :max="formatLength(500, measurementUnit)"
          step="0.1"
          @input="handleWidthInput"
        >
      </label>
      <label class="field">
        <span class="field-label">Height ({{ unitLabel }})</span>
        <input
          v-model="heightCm"
          class="field-input"
          data-testid="image-options-height"
          type="number"
          min="0.1"
          :max="formatLength(500, measurementUnit)"
          step="0.1"
          @input="handleHeightInput"
        >
      </label>
    </div>
    <p
      v-if="error()"
      class="dialog-error"
      data-testid="image-options-size-error"
    >
      {{ error() }}
    </p>

    <label class="field-inline">
      <input
        v-model="lockAspectRatio"
        type="checkbox"
        data-testid="image-options-lock-aspect"
      >
      <span>Lock aspect ratio</span>
    </label>

    <button
      type="button"
      class="reset-button"
      data-testid="image-options-reset-size"
      @click="resetSize"
    >
      Reset size
    </button>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="image-options-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="image-options-apply"
        :disabled="!!error()"
        @click="apply"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* This dialog kept its own scroll cap on the panel (the base's .dialog is
   the short-dialog variant without one). */
:deep(.dialog) {
  max-height: 85vh;
  overflow-y: auto;
}

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

.field-inline {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
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

.dialog-error {
  margin: calc(-1 * var(--space-3)) 0 var(--space-4);
  color: var(--text-danger);
  font-size: var(--text-md);
}

.reset-button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-3) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
  margin-bottom: var(--space-3);
}

</style>
