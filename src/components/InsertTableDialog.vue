<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { MAX_TABLE_COLS, MAX_TABLE_ROWS } from "../lib/core/tables";

// The "Insert table…" dialog (SPEC-022 AC-2): row/column counts beyond the
// grid picker's 10×10. OK emits the size for the parent to insert at the
// caret; Cancel just closes — the draft state lives here and dies here.
// Same structure as CustomSpacingDialog (backdrop, role="dialog", testids).
const emit = defineEmits(["apply", "cancel"]);

const rows = ref("3");
const cols = ref("3");

const rowsInput = useTemplateRef("rowsInput");

onMounted(async () => {
  await nextTick();
  rowsInput.value?.focus();
  rowsInput.value?.select();
});

const error = computed(() => {
  const rowCount = Number(rows.value);
  const colCount = Number(cols.value);
  if (!Number.isInteger(rowCount) || !Number.isInteger(colCount)) {
    return "Enter whole numbers";
  }
  if (rowCount < 1 || rowCount > MAX_TABLE_ROWS) {
    return `Rows must be between 1 and ${MAX_TABLE_ROWS}`;
  }
  if (colCount < 1 || colCount > MAX_TABLE_COLS) {
    return `Columns must be between 1 and ${MAX_TABLE_COLS}`;
  }
  return "";
});

function apply() {
  if (error.value) {
    return;
  }
  emit("apply", { rows: Number(rows.value), cols: Number(cols.value) });
}
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="insert-table-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Insert table"
    >
      <h2 class="dialog-title">
        Insert table
      </h2>
      <div class="field-row">
        <label class="field">
          <span class="field-label">Rows</span>
          <input
            ref="rowsInput"
            v-model="rows"
            class="field-input"
            data-testid="insert-table-rows"
            type="number"
            min="1"
            :max="MAX_TABLE_ROWS"
            step="1"
            @keydown.enter="apply"
          >
        </label>
        <label class="field">
          <span class="field-label">Columns</span>
          <input
            v-model="cols"
            class="field-input"
            data-testid="insert-table-cols"
            type="number"
            min="1"
            :max="MAX_TABLE_COLS"
            step="1"
            @keydown.enter="apply"
          >
        </label>
      </div>
      <p
        v-if="error"
        class="dialog-error"
        data-testid="insert-table-error"
      >
        {{ error }}
      </p>
      <div class="dialog-actions">
        <button
          type="button"
          data-testid="insert-table-cancel"
          @click="emit('cancel')"
        >
          Cancel
        </button>
        <button
          type="button"
          class="primary"
          data-testid="insert-table-confirm"
          :disabled="!!error"
          @click="apply"
        >
          Insert
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
  min-width: 18rem;
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
  font-size: 0.8125rem;
  color: #c5221f;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1rem;
}

.dialog-actions button {
  font: inherit;
  font-size: 0.875rem;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #1a73e8;
  border-color: #1a73e8;
  color: #fff;
}

.dialog-actions button.primary:disabled {
  background: #9e9e9e;
  border-color: #9e9e9e;
  cursor: default;
}
</style>
