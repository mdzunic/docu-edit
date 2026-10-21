<script setup lang="ts">
// The sheet chart's editing surface (SPEC-132, modelled on SlideChartLayer):
// `sheet-chart-toolbar` (role toolbar, "Chart") floating at the top of the
// sheet panel while a chart is selected, hosting ChartToolbar (its reason
// from a read-only chart) and the first-edit `chart-rewrite-notice`; and the
// three dialogs — Select Data, SPEC-052's data sheet and Chart options. On a
// chart sheet Duplicate and Delete are hidden: its chart is never duplicated
// or deleted as a drawing. All state is useSheetCharts'. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetCharts } from "../../views/spreadsheet/useSheetCharts";
import ChartDataSheetDialog from "../ChartDataSheetDialog.vue";
import ChartOptionsDialog from "../ChartOptionsDialog.vue";
import ChartToolbar from "../ChartToolbar.vue";
import SheetChartDataDialog from "./SheetChartDataDialog.vue";

const props = defineProps({
  /** useSheetCharts' return. */
  charts: { type: Object, required: true },
});

const c = computed(() => asSheetCharts(props.charts));
const open = computed(() => c.value.dialog.value);
</script>

<template>
  <div
    v-if="c.selectedChart.value"
    class="sheet-chart-toolbar"
    :class="{ 'is-chart-sheet': c.selectedChart.value.chartSheet }"
    role="toolbar"
    aria-label="Chart"
    data-testid="sheet-chart-toolbar"
  >
    <ChartToolbar
      :unsupported="c.locked.value"
      :reason="c.reason.value"
      @edit-data="c.openData()"
      @options="c.openOptions()"
      @duplicate="c.duplicateChart()"
      @delete="c.deleteChart()"
    />
    <span
      v-if="c.rewriteNotice.value"
      class="chart-rewrite-notice"
      data-testid="chart-rewrite-notice"
    >{{ c.rewriteNotice.value }}</span>
  </div>
  <SheetChartDataDialog
    v-if="open?.kind === 'source'"
    :range="open.range ?? undefined"
    :flags="open.flags"
    :error="c.sourceError.value"
    :picking="c.picking.value"
    :selection-text="c.selectionText.value"
    @apply="(range, flags) => c.applySource(range, flags)"
    @cancel="c.closeDialogs()"
    @pick="c.startPick()"
    @pick-done="c.endPick()"
  />
  <ChartDataSheetDialog
    v-if="open?.kind === 'dataSheet'"
    :initial="open.initial"
    @apply="(data) => c.applyData(data.categories, data.series)"
    @cancel="c.closeDialogs()"
  />
  <ChartOptionsDialog
    v-if="open?.kind === 'options'"
    :initial="open.initial"
    :theme="c.theme.value"
    @apply="(options) => c.applyOptions(options)"
    @series-color="(index, color) => c.applySeriesColor(index, color)"
    @cancel="c.closeDialogs()"
  />
</template>

<style scoped>
.sheet-chart-toolbar {
  position: absolute;
  top: var(--space-2);
  right: var(--space-4);
  z-index: 30;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
  background: var(--surface-base);
  box-shadow: var(--shadow-2);
}

.sheet-chart-toolbar.is-chart-sheet :deep([data-testid="chart-duplicate-button"]),
.sheet-chart-toolbar.is-chart-sheet :deep([data-testid="chart-delete-button"]),
.sheet-chart-toolbar.is-chart-sheet :deep(.chart-toolbar-separator) {
  display: none;
}

.chart-rewrite-notice {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}
</style>
