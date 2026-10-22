<script setup lang="ts">
// Select Data (SPEC-132 "Select Data"): `sheet-chart-data` on BaseDialog —
// the range (`sheet-chart-data-range`, typed, or pointed at through "Select
// cells" `sheet-chart-data-pick`, which collapses the dialog to
// `sheet-chart-data-picking` while the grid's selection fills it; Done
// `sheet-chart-data-pick-done`), series in rows/columns
// (`sheet-chart-data-series-rows` / `-columns`), the first row / first column
// as labels (`sheet-chart-data-first-row` / `-first-col`, mapped onto
// ChartSourceFlags by orientation), `sheet-chart-data-apply` and
// `sheet-chart-data-cancel`. A range too complex to show opens empty with
// `sheet-chart-data-complex` until one is entered. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { COMPLEX_CHART_RANGE_MESSAGE } from "../../lib/spreadsheet/charts";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The prefilled range; null: too complex to show. */
  range: { type: String, default: null },
  /** ChartSourceFlags: { seriesIn, firstRowNames, firstColCategories }. */
  flags: { type: Object, required: true },
  /** The inline error ("" for none). */
  error: { type: String, default: "" },
  /** Collapsed while the grid's selection is pointed at. */
  picking: { type: Boolean, default: false },
  /** The grid's selection as a reference. */
  selectionText: { type: String, default: "" },
});

const emit = defineEmits(["apply", "cancel", "pick", "pick-done"]);

const text = ref(props.range ?? "");
const complex = props.range === null;
const columns = ref(props.flags.seriesIn !== "rows");
const firstRow = ref(columns.value ? !!props.flags.firstRowNames : !!props.flags.firstColCategories);
const firstCol = ref(columns.value ? !!props.flags.firstColCategories : !!props.flags.firstRowNames);
const showComplex = computed(() => complex && text.value.trim() === "");

function apply() {
  emit("apply", text.value, {
    seriesIn: columns.value ? "columns" : "rows",
    firstRowNames: columns.value ? firstRow.value : firstCol.value,
    firstColCategories: columns.value ? firstCol.value : firstRow.value,
  });
}

function pickDone() {
  if (props.selectionText) {
    text.value = props.selectionText;
  }
  emit("pick-done");
}
</script>

<template>
  <div
    v-if="props.picking"
    class="sheet-chart-data-picking"
    role="region"
    aria-label="Select the chart's data"
    data-testid="sheet-chart-data-picking"
  >
    <span>{{ props.selectionText || "Select the chart's data on a sheet" }}</span>
    <UiButton
      data-testid="sheet-chart-data-pick-done"
      @click="pickDone"
    >
      Done
    </UiButton>
  </div>
  <BaseDialog
    v-else
    title="Select data"
    testid="sheet-chart-data"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span>Chart data range</span>
      <span class="row">
        <input
          v-model="text"
          type="text"
          autofocus
          data-testid="sheet-chart-data-range"
          :aria-invalid="props.error ? 'true' : undefined"
          @keydown.enter.prevent="apply"
        >
        <UiButton
          variant="ghost"
          data-testid="sheet-chart-data-pick"
          @click="emit('pick')"
        >
          Select cells
        </UiButton>
      </span>
    </label>
    <p
      v-if="showComplex"
      class="note"
      data-testid="sheet-chart-data-complex"
    >
      {{ COMPLEX_CHART_RANGE_MESSAGE }}
    </p>
    <p
      v-if="props.error"
      class="error"
      role="alert"
      data-testid="sheet-chart-data-error"
    >
      {{ props.error }}
    </p>
    <fieldset class="field">
      <legend>Series in</legend>
      <label><input
        v-model="columns"
        type="radio"
        :value="false"
        data-testid="sheet-chart-data-series-rows"
      > Rows</label>
      <label><input
        v-model="columns"
        type="radio"
        :value="true"
        data-testid="sheet-chart-data-series-columns"
      > Columns</label>
    </fieldset>
    <label><input
      v-model="firstRow"
      type="checkbox"
      data-testid="sheet-chart-data-first-row"
    > First row as {{ columns ? "series names" : "categories" }}</label>
    <label><input
      v-model="firstCol"
      type="checkbox"
      data-testid="sheet-chart-data-first-col"
    > First column as {{ columns ? "categories" : "series names" }}</label>
    <template #actions>
      <UiButton
        variant="ghost"
        data-testid="sheet-chart-data-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        data-testid="sheet-chart-data-apply"
        @click="apply"
      >
        Apply
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  margin: 0 0 var(--space-3);
  border: 0;
  padding: 0;
}

.row {
  display: flex;
  gap: var(--space-2);
}

.note {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.error {
  color: var(--danger);
  font-size: var(--text-sm);
}

.sheet-chart-data-picking {
  position: fixed;
  top: var(--space-4);
  left: 50%;
  z-index: 50;
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  transform: translateX(-50%);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-3);
}
</style>
