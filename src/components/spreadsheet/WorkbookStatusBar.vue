<script setup lang="ts">
// The workbook view's status bar wiring: SpreadsheetStatusBar bound to the
// view-level composables — the selection's aggregates (useSheetSelection),
// the recalculation's status (useRecalc), the cell editor's mode
// (useWorkbookEditing), SPEC-128's filter record count (useSortFilter) and
// the zoom (useSheetZoom, disabled without a grid on show). Extracted so
// SpreadsheetView keeps its line budget (SpreadsheetView.size.spec.ts).
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import SpreadsheetStatusBar from "./SpreadsheetStatusBar.vue";

const props = defineProps({
  /** useWorkbookLoad's return. */
  book: { type: Object, required: true },
  /** useSheetSelection's return. */
  selection: { type: Object, required: true },
  /** useRecalc's return. */
  recalc: { type: Object, required: true },
  /** useWorkbookEditing's return. */
  editing: { type: Object, required: true },
  /** useSheetZoom's return. */
  zoom: { type: Object, required: true },
  /** useSortFilter's return (SPEC-128). */
  sortFilter: { type: Object, default: null },
});

const zoom = computed(() => Object(props.zoom));
const texts = computed(() => ({
  ...Object(props.selection).statusTexts.value,
  ...Object(props.recalc).status.value,
  mode: Object(props.editing).cell.modeLabel.value,
  filter: props.sortFilter ? Object(props.sortFilter).statusText.value : "",
  zoomPercent: zoom.value.percent.value,
  zoomDisabled: Object(props.book).activeSheet.value?.kind !== "grid",
}));
</script>

<template>
  <SpreadsheetStatusBar
    v-bind="texts"
    @zoom-in="zoom.zoomIn()"
    @zoom-out="zoom.zoomOut()"
    @apply-zoom="zoom.applyInput"
    @set-zoom="zoom.set"
  />
</template>
