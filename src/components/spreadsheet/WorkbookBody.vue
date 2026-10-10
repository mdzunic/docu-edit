<script setup lang="ts">
// The workbook view's body (SPEC-116's layout, SPEC-117's surfaces): the
// read-only formula bar (its text from SPEC-119's useFormulaDisplay) above
// the sheet panel — the grid (in Show formulas mode or not), or a foreign
// sheet's placeholder — and the sheet tabs below it, with SPEC-121's tab
// management: the tab context menu, the delete and unhide dialogs and the tab
// colour picker. It only lays the surfaces out and binds them to the
// view-level composables the view hands in (`book` = useWorkbookLoad, `grid`
// = the SheetGridContext of the sheet on show, `editing` =
// useWorkbookEditing); SpreadsheetView keeps only the wiring. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetGridContext } from "../../views/spreadsheet/useSheetGrid";
import { asWorkbookBook } from "../../views/spreadsheet/useWorkbookLoad";
import { asWorkbookEditing } from "../../views/spreadsheet/useWorkbookEditing";
import FormulaBar from "./FormulaBar.vue";
import SheetGrid from "./SheetGrid.vue";
import SheetTabColorPicker from "./SheetTabColorPicker.vue";
import SheetTabDeleteDialog from "./SheetTabDeleteDialog.vue";
import SheetTabMenu from "./SheetTabMenu.vue";
import SheetTabs from "./SheetTabs.vue";
import SheetUnhideDialog from "./SheetUnhideDialog.vue";

const props = defineProps({
  /** useWorkbookLoad's return. */
  book: { type: Object, required: true },
  /** The SheetGridContext of the sheet on show. */
  grid: { type: Object, required: true },
  /** useWorkbookEditing's return (SPEC-121). */
  editing: { type: Object, required: true },
});

const book = computed(() => asWorkbookBook(props.book));
const grid = computed(() => asSheetGridContext(props.grid));
const commands = computed(() => asWorkbookEditing(props.editing).commands);
const panelId = "sheet-panel";

/** The visible tabs with their tab colours resolved for CSS. */
const tabs = computed(() => {
  const sheets = book.value.workbook.value?.sheets ?? [];
  return book.value.tabs.value.map((tab) => ({ ...tab, color: commands.value.tabColorCss(sheets[tab.index]?.tabColor) }));
});
</script>

<template>
  <div
    class="workbook"
    data-testid="spreadsheet-workspace"
  >
    <FormulaBar
      :address="grid.selection.nameBoxText.value"
      :content="grid.formulas.barText.value"
      :spill-member="grid.formulas.spillMember.value"
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
        :show-formulas="grid.formulas.showFormulas.value"
        @cycle-sheet="book.cycleSheet"
      />
    </div>
    <SheetTabs
      :tabs="tabs"
      :active-index="book.activeIndex.value"
      :panel-id="panelId"
      :renaming-index="commands.renaming.value?.index ?? -1"
      :rename-draft="commands.renaming.value?.draft ?? ''"
      :rename-error="commands.renaming.value?.error ?? ''"
      :announcement="commands.announcement.value"
      @select="book.selectSheet"
      @add="commands.insert()"
      @menu="commands.openMenu"
      @rename-start="commands.startRename"
      @rename-input="commands.renameInput"
      @rename-commit="commands.commitRename()"
      @rename-cancel="commands.cancelRename()"
      @rename-blur="commands.blurRename()"
      @move="commands.dropTab"
    />
    <SheetTabMenu
      v-if="commands.menu.value"
      :items="commands.menuItems.value"
      :x="commands.menu.value.x"
      :y="commands.menu.value.y"
      @run="commands.runMenu"
      @close="commands.closeMenu()"
    />
    <SheetTabDeleteDialog
      v-if="commands.deleting.value"
      :name="commands.deleting.value.name"
      @confirm="commands.confirmDelete()"
      @cancel="commands.cancelDelete()"
    />
    <SheetUnhideDialog
      v-if="commands.unhiding.value"
      :sheets="commands.hiddenSheets.value"
      @confirm="commands.confirmUnhide"
      @cancel="commands.cancelUnhide()"
    />
    <SheetTabColorPicker
      v-if="commands.coloring.value"
      :x="commands.coloring.value.x"
      :y="commands.coloring.value.y"
      :theme="commands.theme.value"
      :recent="commands.recentColors.value"
      @select="commands.applyColor"
      @clear="commands.applyColor(null)"
      @close="commands.closeColor()"
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
