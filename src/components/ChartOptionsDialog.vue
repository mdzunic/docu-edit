<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

// The chart options panel (SPEC-052 AC-14..AC-20): title, axis titles,
// legend position, data labels, gridlines, per-series colour and axis
// minimum/maximum with an auto default. Same dialog shape as
// ImageOptionsDialog.vue — a captured `initial` prefill, apply/cancel
// emits, local draft refs, OK disabled while invalid. BaseDialog owns the
// shell, Escape and initial focus (the [autofocus] title input); @dismiss
// maps onto cancel.
const props = defineProps({
  /** Prefill: { title, categoryAxisTitle, valueAxisTitle, legend,
   * dataLabels, gridlines, valueAxisMin, valueAxisMax, series: [{name,
   * color}], hasAxes: boolean }. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "series-color", "cancel"]);

const title = ref(String(props.initial.title || ""));
const categoryAxisTitle = ref(String(props.initial.categoryAxisTitle || ""));
const valueAxisTitle = ref(String(props.initial.valueAxisTitle || ""));
const legend = ref(String(props.initial.legend || "none"));
const dataLabels = ref(!!props.initial.dataLabels);
const gridlines = ref(!!props.initial.gridlines);
const autoMin = ref(props.initial.valueAxisMin === undefined || props.initial.valueAxisMin === null);
const autoMax = ref(props.initial.valueAxisMax === undefined || props.initial.valueAxisMax === null);
const valueAxisMin = ref(props.initial.valueAxisMin !== undefined && props.initial.valueAxisMin !== null ? String(props.initial.valueAxisMin) : "0");
const valueAxisMax = ref(props.initial.valueAxisMax !== undefined && props.initial.valueAxisMax !== null ? String(props.initial.valueAxisMax) : "100");
const hasAxes = !!props.initial.hasAxes;
const series = ref((props.initial.series || []).map((s = { name: "", color: "#000000" }) => ({ name: s.name, color: s.color })));

function minMaxError() {
  if (!autoMin.value && !Number.isFinite(Number(valueAxisMin.value))) {
    return "Axis minimum must be a number";
  }
  if (!autoMax.value && !Number.isFinite(Number(valueAxisMax.value))) {
    return "Axis maximum must be a number";
  }
  if (!autoMin.value && !autoMax.value && Number(valueAxisMin.value) >= Number(valueAxisMax.value)) {
    return "Axis minimum must be less than the maximum";
  }
  return "";
}

function onSeriesColorInput(index = 0, event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    series.value[index].color = target.value;
    emit("series-color", index, target.value);
  }
}

function apply() {
  if (minMaxError()) {
    return;
  }
  emit("apply", {
    title: title.value.trim() || undefined,
    categoryAxisTitle: categoryAxisTitle.value.trim() || undefined,
    valueAxisTitle: valueAxisTitle.value.trim() || undefined,
    legend: legend.value,
    dataLabels: dataLabels.value,
    gridlines: gridlines.value,
    valueAxisMin: autoMin.value ? undefined : Number(valueAxisMin.value),
    valueAxisMax: autoMax.value ? undefined : Number(valueAxisMax.value),
  });
}
</script>

<template>
  <BaseDialog
    title="Chart options"
    testid="chart-options-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span class="field-label">Title</span>
      <input
        v-model="title"
        class="field-input"
        data-testid="chart-options-title"
        type="text"
        autofocus
      >
    </label>

    <div
      v-if="hasAxes"
      class="field-row"
    >
      <label class="field">
        <span class="field-label">Category axis title</span>
        <input
          v-model="categoryAxisTitle"
          class="field-input"
          data-testid="chart-options-category-axis-title"
          type="text"
        >
      </label>
      <label class="field">
        <span class="field-label">Value axis title</span>
        <input
          v-model="valueAxisTitle"
          class="field-input"
          data-testid="chart-options-value-axis-title"
          type="text"
        >
      </label>
    </div>

    <label class="field">
      <span class="field-label">Legend position</span>
      <select
        v-model="legend"
        class="field-input"
        data-testid="chart-options-legend"
      >
        <option value="none">
          None
        </option>
        <option value="top">
          Top
        </option>
        <option value="right">
          Right
        </option>
        <option value="bottom">
          Bottom
        </option>
        <option value="left">
          Left
        </option>
      </select>
    </label>

    <label class="field-inline">
      <input
        v-model="dataLabels"
        type="checkbox"
        data-testid="chart-options-data-labels"
      >
      <span>Data labels</span>
    </label>

    <label
      v-if="hasAxes"
      class="field-inline"
    >
      <input
        v-model="gridlines"
        type="checkbox"
        data-testid="chart-options-gridlines"
      >
      <span>Gridlines</span>
    </label>

    <div
      v-if="hasAxes"
      class="field-row"
    >
      <label class="field-inline">
        <input
          v-model="autoMin"
          type="checkbox"
          data-testid="chart-options-auto-min"
        >
        <span>Auto minimum</span>
      </label>
      <input
        v-model="valueAxisMin"
        class="field-input"
        data-testid="chart-options-value-min"
        type="number"
        step="any"
        :disabled="autoMin"
      >
      <label class="field-inline">
        <input
          v-model="autoMax"
          type="checkbox"
          data-testid="chart-options-auto-max"
        >
        <span>Auto maximum</span>
      </label>
      <input
        v-model="valueAxisMax"
        class="field-input"
        data-testid="chart-options-value-max"
        type="number"
        step="any"
        :disabled="autoMax"
      >
    </div>
    <p
      v-if="minMaxError()"
      class="dialog-error"
      data-testid="chart-options-minmax-error"
    >
      {{ minMaxError() }}
    </p>

    <div class="field">
      <span class="field-label">Series colours</span>
      <label
        v-for="(s, index) in series"
        :key="index"
        class="field-inline"
      >
        <input
          type="color"
          :data-testid="`chart-options-series-color-${index}`"
          :value="s.color"
          @input="onSeriesColorInput(Number(index), $event)"
        >
        <span>{{ s.name }}</span>
      </label>
    </div>

    <template #actions>
      <button
        type="button"
        data-testid="chart-options-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="chart-options-apply"
        :disabled="!!minMaxError()"
        @click="apply"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* This dialog's body can be tall (axis rows + a colour row per series), so
   the panel keeps its own scroll cap. */
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
  align-items: center;
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

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
