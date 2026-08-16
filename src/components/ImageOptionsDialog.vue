<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

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
});

const emit = defineEmits(["apply", "cancel"]);

const altText = ref(String(props.initial.altText || ""));
const widthCm = ref(String(props.initial.widthCm));
const heightCm = ref(String(props.initial.heightCm));
const lockAspectRatio = ref(props.initial.lockAspectRatio !== false);

const naturalWidthCm = Number(props.initial.naturalWidthCm) || Number(props.initial.widthCm) || 1;
const naturalHeightCm = Number(props.initial.naturalHeightCm) || Number(props.initial.heightCm) || 1;
// The ratio lock preserves — SPEC-024 AC-19's "original aspect ratio" is the
// image's natural (source) ratio, not whatever the box currently shows.
const naturalRatio = naturalWidthCm / naturalHeightCm;

function widthError() {
  const value = Number(widthCm.value);
  return Number.isFinite(value) && value > 0 && value <= 500 ? "" : "Width must be between 0 and 500cm";
}

function heightError() {
  const value = Number(heightCm.value);
  return Number.isFinite(value) && value > 0 && value <= 500 ? "" : "Height must be between 0 and 500cm";
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
  const width = Number(widthCm.value);
  if (Number.isFinite(width) && width > 0) {
    heightCm.value = (width / naturalRatio).toFixed(2);
  }
}

function handleHeightInput() {
  if (!lockAspectRatio.value) {
    return;
  }
  const height = Number(heightCm.value);
  if (Number.isFinite(height) && height > 0) {
    widthCm.value = (height * naturalRatio).toFixed(2);
  }
}

// AC-21, AC-22: restores natural size and re-enables the lock.
function resetSize() {
  widthCm.value = naturalWidthCm.toFixed(2);
  heightCm.value = naturalHeightCm.toFixed(2);
  lockAspectRatio.value = true;
}

function apply() {
  if (error()) {
    return;
  }
  emit("apply", {
    altText: altText.value.trim(),
    widthCm: Number(widthCm.value),
    heightCm: Number(heightCm.value),
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
        <span class="field-label">Width (cm)</span>
        <input
          v-model="widthCm"
          class="field-input"
          data-testid="image-options-width"
          type="number"
          min="0.1"
          max="500"
          step="0.1"
          @input="handleWidthInput"
        >
      </label>
      <label class="field">
        <span class="field-label">Height (cm)</span>
        <input
          v-model="heightCm"
          class="field-input"
          data-testid="image-options-height"
          type="number"
          min="0.1"
          max="500"
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
      <button
        type="button"
        data-testid="image-options-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="image-options-apply"
        :disabled="!!error()"
        @click="apply"
      >
        OK
      </button>
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
  gap: 0.25rem;
  margin-bottom: 0.75rem;
}

.field-row {
  display: flex;
  gap: 0.75rem;
}

.field-row .field {
  flex: 1;
}

.field-inline {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.75rem;
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

.dialog-error {
  margin: -0.5rem 0 0.75rem;
  color: #b3261e;
  font-size: 0.8125rem;
}

.reset-button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.375rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  margin-bottom: 0.5rem;
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
