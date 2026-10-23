<script setup lang="ts">
// The workbook view's not-found state (SPEC-116): a routed id that names no
// workbook in the library. Extracted from SpreadsheetView (SPEC-121) so the
// view stays a wiring hub within SpreadsheetView.size.spec.ts's budget.
// SPEC-137: while a stored workbook is read back from browser storage
// (useWorkbookLoad provides the flag), it shows the loading state instead —
// never "doesn't exist".
import { inject, ref } from "vue";
import { WORKBOOK_LOADING_KEY } from "../../views/spreadsheet/useWorkbookLoad";

const loading = inject(WORKBOOK_LOADING_KEY, ref(false));
</script>

<template>
  <div
    v-if="loading"
    class="workbook-missing"
    data-testid="spreadsheet-loading"
    role="status"
  >
    <p>Opening workbook…</p>
  </div>
  <div
    v-else
    class="workbook-missing"
    data-testid="spreadsheet-not-found"
    role="alert"
  >
    <p>This workbook doesn't exist in your library.</p>
  </div>
</template>

<style scoped>
.workbook-missing {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--text-secondary);
}
</style>
