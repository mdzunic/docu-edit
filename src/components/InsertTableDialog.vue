<script setup lang="ts">
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { MAX_TABLE_COLS, MAX_TABLE_ROWS } from "../lib/core/tables";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

// The "Insert table…" dialog (SPEC-022 AC-2): row/column counts beyond the
// grid picker's 10×10. OK emits the size for the parent to insert at the
// caret; Cancel just closes — the draft state lives here and dies here.
// BaseDialog owns the shell, Escape and initial focus (the [autofocus]
// rows input).
const emit = defineEmits(["apply", "cancel"]);

const rows = ref("3");
const cols = ref("3");
// SPEC-076 AC-20: checked by default — SPEC-043's accessibility checker
// flags a table with no header row, so a table made here passes that check
// unless the author says otherwise. The Insert menu's grid quick-pick has no
// dialog and is unchanged: it still creates no header row (AC-21).
const headerRow = ref(true);

const rowsInput = useTemplateRef("rowsInput");

// BaseDialog focuses the [autofocus] rows input; selecting the prefilled
// count on top of that (type-to-replace) is this dialog's own one-shot
// behaviour.
onMounted(async () => {
  await nextTick();
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
  emit("apply", { rows: Number(rows.value), cols: Number(cols.value), headerRow: headerRow.value });
}
</script>

<template>
  <BaseDialog
    title="Insert table"
    testid="insert-table-dialog"
    width="18rem"
    @dismiss="emit('cancel')"
  >
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
          autofocus
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
    <label class="field-check">
      <input
        v-model="headerRow"
        type="checkbox"
        data-testid="insert-table-header-row"
      >
      <span>First row is a header row</span>
    </label>
    <p
      v-if="error"
      class="dialog-error"
      data-testid="insert-table-error"
    >
      {{ error }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="insert-table-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="insert-table-confirm"
        :disabled="!!error"
        @click="apply"
      >
        Insert
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field-row {
  display: flex;
  gap: var(--space-4);
  margin-bottom: var(--space-4);
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
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

.field-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  font-size: var(--text-md);
  color: var(--danger);
}
</style>
