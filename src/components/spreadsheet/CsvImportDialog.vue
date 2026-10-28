<script setup lang="ts">
// SPEC-126's "Text import options" dialog (`csv-import-dialog`): Separator,
// Character set and "Numbers and dates" selects, preselected from the
// workbook's recorded dialect, and a live preview of the first records
// (`csv-import-preview`) re-parsed from the retained source bytes on every
// change. The primary action is "Apply" (`csv-import-apply`) while the
// workbook is unedited and "Open as new workbook" (`csv-import-open-new`)
// once it is; useCsvImport runs both. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asCsvImport, CSV_ENCODING_OPTIONS, CSV_LOCALE_OPTIONS, CSV_SEPARATOR_OPTIONS } from "../../views/spreadsheet/useCsvImport";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import UiSelect from "../ui/UiSelect.vue";

const props = defineProps({
  /** useCsvImport's return. */
  csvImport: { type: Object, required: true },
});

const csv = computed(() => asCsvImport(props.csvImport));
const separators = CSV_SEPARATOR_OPTIONS.map((option) => ({ value: option.value, label: option.label }));
const encodings = CSV_ENCODING_OPTIONS.map((option) => ({ value: option.value, label: option.label }));
const locales = CSV_LOCALE_OPTIONS.slice();
</script>

<template>
  <BaseDialog
    v-if="csv.dialogOpen.value"
    title="Text import options"
    testid="csv-import-dialog"
    width="34rem"
    @dismiss="csv.closeDialog()"
  >
    <div class="csv-import-fields">
      <label class="csv-import-field">
        <span>Separator</span>
        <UiSelect
          v-model="csv.separator.value"
          :options="separators"
          data-testid="csv-import-separator"
        />
      </label>
      <label class="csv-import-field">
        <span>Character set</span>
        <UiSelect
          v-model="csv.encoding.value"
          :options="encodings"
          data-testid="csv-import-encoding"
        />
      </label>
      <label class="csv-import-field">
        <span>Numbers and dates</span>
        <UiSelect
          v-model="csv.localeTag.value"
          :options="locales"
          data-testid="csv-import-locale"
        />
      </label>
    </div>
    <div
      class="csv-import-preview"
      data-testid="csv-import-preview"
      aria-label="Preview"
    >
      <p
        v-if="csv.preview.value.error"
        class="csv-import-preview-error"
      >
        {{ csv.preview.value.error }}
      </p>
      <table v-else>
        <tbody>
          <tr
            v-for="(record, row) in csv.preview.value.records"
            :key="row"
            data-testid="csv-import-preview-row"
          >
            <td
              v-for="(field, col) in record"
              :key="col"
            >
              {{ field }}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <p
      v-if="csv.error.value"
      class="csv-import-error"
      role="alert"
      data-testid="csv-import-error"
    >
      {{ csv.error.value }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="csv-import-cancel"
        @click="csv.closeDialog()"
      >
        Cancel
      </UiButton>
      <UiButton
        v-if="csv.edited.value"
        variant="primary"
        data-testid="csv-import-open-new"
        @click="csv.openAsNew()"
      >
        Open as new workbook
      </UiButton>
      <UiButton
        v-else
        variant="primary"
        data-testid="csv-import-apply"
        :disabled="csv.locked.value"
        @click="csv.apply()"
      >
        Apply
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.csv-import-fields {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: var(--space-3);
}

.csv-import-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.csv-import-preview {
  max-height: 16rem;
  margin-top: var(--space-3);
  overflow: auto;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  font-family: var(--font-mono, monospace);
  font-size: var(--text-xs);
}

.csv-import-preview table {
  border-collapse: collapse;
}

.csv-import-preview td {
  padding: var(--space-1) var(--space-2);
  border-right: var(--border-width) solid var(--border-subtle, var(--border-default));
  border-bottom: var(--border-width) solid var(--border-subtle, var(--border-default));
  white-space: pre;
}

.csv-import-preview-error,
.csv-import-error {
  margin: var(--space-2) 0 0;
  color: var(--danger);
  font-size: var(--text-sm);
}

.csv-import-preview-error {
  padding: var(--space-2);
}
</style>
