<script setup lang="ts">
// The workbook's version history pane (SPEC-138): the body of SPEC-129's sheet
// side panel when its `"history"` pane is active — the product-neutral
// HistoryPanel (`sheet-history-panel`, SPEC-103's `history-*` testids) over
// useWorkbookHistory. A preview shows in the grid itself, so the banner here
// carries no body of its own: "Previewing an earlier version, read-only" and
// "Back to current". Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import HistoryPanel from "../HistoryPanel.vue";
import { WORKBOOK_RESTORE_NOTE } from "../../views/spreadsheet/workbookHistoryLabels";

defineProps({
  /** useWorkbookHistory's return. */
  controller: { type: Object, required: true },
});
</script>

<template>
  <HistoryPanel
    class="workbook-history"
    :controller="controller"
    testid="sheet-history-panel"
    label="Version history"
    :restore-note="WORKBOOK_RESTORE_NOTE"
  />
</template>

<style scoped>
/* The pane's whole height, scrolling within it. */
.workbook-history {
  height: 100%;
  min-height: 0;
}
</style>
