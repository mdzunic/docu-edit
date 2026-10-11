<script setup lang="ts">
// SPEC-126's import bar (`csv-import-bar`) on a csv/tsv workbook: how the
// file was read — "Opened as semicolon-separated, Windows-1252, numbers and
// dates read as German · Change…" — with Change… (`csv-import-change`)
// opening the "Text import options" dialog and a dismiss button
// (`csv-import-dismiss`). It renders only while useCsvImport has a sentence
// for it (a csv/tsv workbook, not dismissed). Runtime prop declarations per
// the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asCsvImport } from "../../views/spreadsheet/useCsvImport";
import UiIconButton from "../ui/UiIconButton.vue";

const props = defineProps({
  /** useCsvImport's return. */
  csvImport: { type: Object, required: true },
});

const csv = computed(() => asCsvImport(props.csvImport));
</script>

<template>
  <div
    v-if="csv.barText.value"
    class="csv-import-bar"
    role="status"
    data-testid="csv-import-bar"
  >
    <span class="csv-import-text">{{ `${csv.barText.value} · ` }}</span>
    <button
      type="button"
      class="csv-import-change"
      data-testid="csv-import-change"
      @click="csv.openDialog()"
    >
      Change…
    </button>
    <UiIconButton
      class="csv-import-dismiss"
      icon="close"
      label="Dismiss"
      data-testid="csv-import-dismiss"
      @click="csv.dismiss()"
    />
  </div>
</template>

<style scoped>
.csv-import-bar {
  display: flex;
  align-items: center;
  gap: var(--space-1);
  padding: var(--space-2) var(--space-5);
  background: var(--surface-selected);
  border-bottom: var(--border-width) solid var(--accent-subtle);
  color: var(--text-accent);
  font-size: var(--text-md);
}

.csv-import-change {
  padding: 0;
  border: 0;
  background: transparent;
  color: inherit;
  font: inherit;
  font-weight: 600;
  text-decoration: underline;
  cursor: pointer;
}

.csv-import-change:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.csv-import-dismiss {
  margin-left: auto;
}
</style>
