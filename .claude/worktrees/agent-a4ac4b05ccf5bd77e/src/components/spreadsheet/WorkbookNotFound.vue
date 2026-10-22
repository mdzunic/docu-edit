<script setup lang="ts">
// The workbook view's not-found state (SPEC-116): a routed id that names no
// workbook in the library. Extracted from SpreadsheetView (SPEC-121) so the
// view stays a wiring hub within SpreadsheetView.size.spec.ts's budget.
</script>

<template>
  <div
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
