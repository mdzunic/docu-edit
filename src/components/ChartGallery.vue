<script setup lang="ts">
// "Insert -> Chart"'s gallery (SPEC-052 AC-2): the closed chart catalogue.
// Picking one inserts a chart with sample data at the caret (AC-3) — this
// component only reports the pick, DocumentView.vue's useCharts owns what
// happens next.
//
// Each type shows as a sketch of itself (chartIcons.ts, which borrows the
// painter's own bar spacing, doughnut hole and axis rule) rather than its
// name: same-size cells line the columns up with each other, with
// ShapeGallery.vue's grid, and with the menu's 1rem text gutter. The names
// stay as each button's tooltip and accessible name.
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
import { asChartType, CHART_CATALOGUE } from "../lib/core/charts";
import { CHART_ICON_SIZE, chartIconMarks } from "../lib/core/chartIcons";

defineEmits(["pick"]);

const ICON_VIEWBOX = `0 0 ${CHART_ICON_SIZE} ${CHART_ICON_SIZE}`;

function iconMarks(chartType = "columnClustered") {
  return chartIconMarks(asChartType(chartType));
}
</script>

<template>
  <div
    class="chart-gallery"
    data-testid="chart-gallery"
    role="menu"
    aria-label="Chart gallery"
  >
    <div class="chart-gallery-items">
      <button
        v-for="item in CHART_CATALOGUE"
        :key="item.chartType"
        type="button"
        class="chart-gallery-item"
        :data-testid="`chart-gallery-${item.chartType}`"
        :title="item.label"
        :aria-label="item.label"
        @mousedown.prevent
        @click="$emit('pick', item.chartType)"
      >
        <svg
          class="chart-gallery-icon"
          :viewBox="ICON_VIEWBOX"
          aria-hidden="true"
          focusable="false"
        >
          <path
            v-for="(mark, index) in iconMarks(item.chartType)"
            :key="index"
            :class="`chart-gallery-mark-${mark.role}`"
            :d="mark.d"
          />
        </svg>
      </button>
    </div>
  </div>
</template>

<style scoped>
.chart-gallery {
  /* The 1rem side padding is MenuBar's own command gutter, so the grid
     starts and ends on the same vertical lines as the menu's labels above
     and below it — and as ShapeGallery.vue's own grid. */
  padding: 0.125rem 1rem 0.25rem;
}

.chart-gallery-items {
  display: grid;
  /* One column per catalogue entry, however many that is: the whole gallery
     is a single row, so it costs the Insert menu — the longest in the bar,
     and one that has to fit a 720px viewport whole (SPEC-028 AC-8/AC-9) —
     one row of height rather than two. */
  grid-auto-flow: column;
  grid-auto-columns: 1fr;
  gap: 0.25rem;
}

.chart-gallery-item {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 1.5rem;
  padding: 0.125rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.chart-gallery-item:hover,
.chart-gallery-item:focus-visible {
  background: #f1f3f4;
  border-color: #1a73e8;
  outline: none;
}

.chart-gallery-icon {
  display: block;
  width: 1.25rem;
  height: 1.25rem;
}

.chart-gallery-mark-axis {
  fill: none;
  stroke: #9aa0a6;
  stroke-width: 1;
}

.chart-gallery-mark-fill {
  fill: #e8f0fe;
  fill-rule: evenodd;
  stroke: #3c4043;
  stroke-width: 1.25;
  stroke-linejoin: round;
}

.chart-gallery-mark-line {
  fill: none;
  stroke: #3c4043;
  stroke-width: 1.5;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.chart-gallery-mark-dot {
  fill: #3c4043;
  stroke: none;
}
</style>
