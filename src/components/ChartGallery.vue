<script setup lang="ts">
// "Insert -> Chart"'s gallery (SPEC-052 AC-2): the closed chart catalogue.
// Picking one inserts a chart with sample data at the caret (AC-3) — this
// component only reports the pick, DocumentView.vue's useCharts owns what
// happens next. Mirrors ShapeGallery.vue's own pattern.
import { CHART_CATALOGUE } from "../lib/documentModel";

defineEmits(["pick"]);
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
        @mousedown.prevent
        @click="$emit('pick', item.chartType)"
      >
        {{ item.label }}
      </button>
    </div>
  </div>
</template>

<style scoped>
.chart-gallery {
  padding: 0.125rem;
  min-width: 220px;
}

.chart-gallery-items {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 0.125rem;
}

.chart-gallery-item {
  font: inherit;
  font-size: 0.6875rem;
  line-height: 1.1;
  padding: 0.25rem 0.3rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  color: #202124;
  cursor: pointer;
  text-align: left;
  white-space: nowrap;
}

.chart-gallery-item:hover {
  background: #f1f3f4;
}
</style>
