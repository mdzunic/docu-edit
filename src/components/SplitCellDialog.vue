<script setup lang="ts">
import { nextTick, onMounted, ref, useTemplateRef } from "vue";

// The "Split cell…" dialog (SPEC-023 AC-8, AC-9): asks for the number of
// rows and columns to split the selected cell into. OK emits the pair for
// the parent to apply; Cancel just closes.
const emit = defineEmits(["apply", "cancel"]);

const rows = ref("1");
const cols = ref("2");

const rowsInput = useTemplateRef("rowsInput");

onMounted(async () => {
  await nextTick();
  rowsInput.value?.focus();
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
  <div
    class="dialog-backdrop"
    data-testid="split-cell-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Split cell"
    >
      <h2 class="dialog-title">
        Split cell
      </h2>
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
      <div class="dialog-actions">
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
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  min-width: 20rem;
  max-width: 90vw;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

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

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1.25rem;
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
