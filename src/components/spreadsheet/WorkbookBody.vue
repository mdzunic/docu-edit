<script setup lang="ts">
// The workbook view's body (SPEC-116's layout, SPEC-117's surfaces): the
// read-only formula bar above the sheet panel — the grid, or a foreign
// sheet's placeholder — and the sheet tabs below it. It only lays the
// surfaces out and binds them to the view-level composables the view hands
// in (`book` = useWorkbookLoad, `grid` = the SheetGridContext of the sheet on
// show); SpreadsheetView keeps only the wiring. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetGridContext } from "../../views/spreadsheet/useSheetGrid";
import { asWorkbookBook } from "../../views/spreadsheet/useWorkbookLoad";
import FormulaBar from "./FormulaBar.vue";
import SheetGrid from "./SheetGrid.vue";
import SheetTabs from "./SheetTabs.vue";

const props = defineProps({
  /** useWorkbookLoad's return. */
  book: { type: Object, required: true },
  /** The SheetGridContext of the sheet on show. */
  grid: { type: Object, required: true },
});

const book = computed(() => asWorkbookBook(props.book));
const grid = computed(() => asSheetGridContext(props.grid));
const panelId = "sheet-panel";
</script>

<template>
  <div
    class="workbook"
    data-testid="spreadsheet-workspace"
  >
    <FormulaBar
      :address="grid.selection.nameBoxText.value"
      :content="grid.selection.activeContent.value"
    />
    <div
      :id="panelId"
      class="workbook-sheet"
      role="tabpanel"
      :aria-labelledby="book.activeSheet.value ? `sheet-tab-${book.activeSheet.value.id}` : undefined"
    >
      <div
        v-if="book.activeSheet.value?.kind === 'foreign'"
        class="sheet-foreign"
        data-testid="sheet-foreign"
      >
        {{ book.foreignMessage.value }}
      </div>
      <SheetGrid
        v-else-if="book.activeSheet.value"
        :key="`${book.activeSheet.value.id}:${book.generation.value}`"
        :sheet="book.activeSheet.value"
        :context="grid"
        @cycle-sheet="book.cycleSheet"
      />
    </div>
    <SheetTabs
      :tabs="book.tabs.value"
      :active-index="book.activeIndex.value"
      :panel-id="panelId"
      @select="book.selectSheet"
    />
  </div>
</template>

<style scoped>
.workbook {
  height: 100%;
  min-height: 0;
  display: flex;
  flex-direction: column;
}

.workbook-sheet {
  flex: 1 1 auto;
  min-height: 0;
}

.sheet-foreign {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--text-secondary);
}
</style>
