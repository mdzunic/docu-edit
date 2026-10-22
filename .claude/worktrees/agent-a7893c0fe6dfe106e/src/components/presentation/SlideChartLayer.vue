<script setup lang="ts">
// The chart editing surface (SPEC-092): the contextual chart toolbar shown for
// a selected chart frame plus SPEC-052's data-sheet and options dialogs, reused
// verbatim (ChartToolbar/ChartDataSheetDialog/ChartOptionsDialog unchanged).
// Extracted from PresentationView so the view keeps only a one-line hookup
// (PresentationView.size.spec.ts). All state and commands live on the
// `useSlideCharts` controller passed in; this component is pure wiring +
// chrome. Runtime prop declarations per the espree constraint.
import ChartToolbar from "../ChartToolbar.vue";
import ChartDataSheetDialog from "../ChartDataSheetDialog.vue";
import ChartOptionsDialog from "../ChartOptionsDialog.vue";

const props = defineProps({
  /** The useSlideCharts controller. */
  controller: { type: Object, required: true },
});
const c = props.controller;
</script>

<template>
  <!-- The floating chart toolbar (AC-14/15): "Edit data"/"Chart options" for a
       supported chart, the disabled controls + reason for an unsupported one,
       and the first-edit rewrite notice (AC-23). -->
  <div
    v-if="c.selectedChart.value"
    class="slide-chart-toolbar"
    role="toolbar"
    aria-label="Chart"
    data-testid="slide-chart-toolbar"
  >
    <ChartToolbar
      :unsupported="c.isUnsupported.value"
      @edit-data="c.openDataSheet()"
      @options="c.openOptions()"
      @delete="c.deleteChart()"
      @duplicate="c.duplicateChart()"
    />
    <span
      v-if="c.rewriteNotice.value"
      class="chart-rewrite-notice"
      data-testid="chart-rewrite-notice"
    >{{ c.rewriteNotice.value }}</span>
  </div>

  <ChartDataSheetDialog
    v-if="c.editingData.value"
    :initial="c.dataSheetInitial.value"
    @apply="(data) => c.applyData(data.categories, data.series)"
    @cancel="c.closeDialogs()"
  />
  <ChartOptionsDialog
    v-if="c.editingOptions.value"
    :initial="c.optionsInitial.value"
    @apply="(options) => c.applyOptions(options)"
    @series-color="(index, color) => c.applySeriesColor(index, color)"
    @cancel="c.closeDialogs()"
  />
</template>

<style scoped>
.slide-chart-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised, var(--surface-default));
}

.chart-rewrite-notice {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}
</style>
