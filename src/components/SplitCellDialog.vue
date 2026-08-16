<script setup lang="ts">
import { nextTick, onMounted, ref, useTemplateRef } from "vue";
import BaseDialog from "./BaseDialog.vue";

// The "Split cell…" dialog (SPEC-023 AC-8, AC-9): asks for the number of
// rows and columns to split the selected cell into. OK emits the pair for
// the parent to apply; Cancel just closes. BaseDialog owns the shell,
// Escape and initial focus (the [autofocus] rows input).
const emit = defineEmits(["apply", "cancel"]);

const rows = ref("1");
const cols = ref("2");

const rowsInput = useTemplateRef("rowsInput");

// BaseDialog focuses the [autofocus] rows input; selecting the prefilled
// count on top of that (type-to-replace) is this dialog's own one-shot
// behaviour.
onMounted(async () => {
  await nextTick();
  rowsInput.value?.select();
});

// A generous but finite ceiling — matches the table's own row/column caps
// closely enough that a mistyped value fails the dialog, not the command.
const MAX_SPLIT = 63;

function parsedCount(raw = "1") {
  const value = Number(raw);
  return Number.isInteger(value) ? value : NaN;
}

function fieldError(raw = "1") {
  const value = parsedCount(raw);
  if (!Number.isInteger(value) || value < 1 || value > MAX_SPLIT) {
    return `Enter a whole number between 1 and ${MAX_SPLIT}`;
  }
  return "";
}

function error() {
  return fieldError(rows.value) || fieldError(cols.value);
}

function apply() {
  if (error()) {
    return;
  }
  emit("apply", { rows: parsedCount(rows.value), cols: parsedCount(cols.value) });
}
</script>

<template>
  <BaseDialog
    title="Split cell"
    testid="split-cell-dialog"
    @dismiss="emit('cancel')"
  >
    <div class="field-row">
      <label class="field">
        <span class="field-label">Number of rows</span>
        <input
          ref="rowsInput"
          v-model="rows"
          class="field-input"
          data-testid="split-cell-rows"
          type="number"
          min="1"
          :max="MAX_SPLIT"
          autofocus
          @keydown.enter="apply"
        >
      </label>
      <label class="field">
        <span class="field-label">Number of columns</span>
        <input
          v-model="cols"
          class="field-input"
          data-testid="split-cell-cols"
          type="number"
          min="1"
          :max="MAX_SPLIT"
          @keydown.enter="apply"
        >
      </label>
    </div>
    <p
      v-if="error()"
      class="dialog-error"
      data-testid="split-cell-error"
    >
      {{ error() }}
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="split-cell-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="split-cell-apply"
        :disabled="!!error()"
        @click="apply"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field-row {
  display: flex;
  gap: 0.75rem;
  margin-bottom: 0.75rem;
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
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
