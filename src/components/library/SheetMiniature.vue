<script setup lang="ts">
// A workbook's library miniature (SPEC-116, card and quick look): the active
// sheet's top-left corner as a tiny grid, built from `model.preview` — up to
// 12 × 6 General-format strings the store keeps on the shell record, so the
// card renders without rejoining a single row. Runtime prop declarations per
// the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetPreview } from "../../views/spreadsheet/useWorkbookLoad";

const props = defineProps({
  /** The SheetPreview to draw (absent for a workbook stored without one). */
  preview: { type: Object, default: undefined },
});

const preview = computed(() => asSheetPreview(props.preview));
const columns = computed(() => Math.max(1, ...preview.value.cells.map((row) => row.length)));
</script>

<template>
  <div
    class="sheet-miniature"
    data-testid="sheet-miniature"
  >
    <div
      class="miniature-grid"
      :style="{ gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))` }"
    >
      <template
        v-for="(row, rowIndex) in preview.cells"
        :key="rowIndex"
      >
        <span
          v-for="(text, colIndex) in row"
          :key="colIndex"
          class="miniature-cell"
          data-testid="sheet-miniature-cell"
        >{{ text }}</span>
      </template>
    </div>
    <span
      v-if="preview.sheetName"
      class="miniature-tab"
    >{{ preview.sheetName }}</span>
  </div>
</template>

<style scoped>
.sheet-miniature {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  background: var(--surface-raised);
}

.miniature-grid {
  display: grid;
  flex: 1 1 auto;
  align-content: start;
  min-height: 0;
  border-left: var(--border-width) solid var(--border-subtle);
}

.miniature-cell {
  min-width: 0;
  height: 0.9rem;
  padding: 0 2px;
  overflow: hidden;
  border-right: var(--border-width) solid var(--border-subtle);
  border-bottom: var(--border-width) solid var(--border-subtle);
  color: var(--text-primary);
  font-size: 0.5rem;
  line-height: 0.9rem;
  white-space: nowrap;
}

.miniature-tab {
  flex: none;
  align-self: flex-start;
  max-width: 100%;
  padding: 0 var(--space-2);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  border-top: var(--border-width) solid var(--border-default);
  color: var(--text-secondary);
  font-size: var(--text-xs);
}
</style>
