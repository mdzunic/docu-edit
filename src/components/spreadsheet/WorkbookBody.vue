<script setup lang="ts">
// The workbook view's body (SPEC-116's layout, SPEC-117's surfaces): the
// formula bar (its text from SPEC-119's useFormulaDisplay; editable on a grid
// sheet since SPEC-122, sharing the in-cell editor's draft) above the sheet
// panel — the grid (in Show formulas mode or not), or a foreign sheet's
// placeholder — and the sheet tabs below it, with SPEC-121's tab management:
// the tab context menu, the delete and unhide dialogs and the tab colour
// picker. A tab chosen while a cell is edited goes through the cell editor
// (SPEC-122: pointing shows that sheet and keeps editing; otherwise the edit
// commits first). It only lays the surfaces out and binds them to the
// view-level composables the view hands in (`book` = useWorkbookLoad, `grid`
// = the SheetGridContext of the sheet on show, `editing` =
// useWorkbookEditing); SpreadsheetView keeps only the wiring. SPEC-123: the
// cell context menu and Format Cells, from the grid context's `formatting`.
// SPEC-124: the header menu, the insert/delete and size dialogs (the grid
// context's `structure`) and the merge warning (`merges`); the cell menu's
// Insert…/Delete… go to the structure. SPEC-125: the cell menu's Cut, Copy,
// Paste and Paste special… (`clipboard`), Paste special's dialog and the
// drag-move's overwrite question (`rangeMove`). SPEC-128: the Sort dialog,
// the filter popover and its dialogs and the refusal alert
// (SortFilterSurfaces, the grid context's `sortFilter`). SPEC-129: the
// sheet's side panel beside the sheet (SheetSidePanel, from `editing`).
// SPEC-130: the Data validation dialog (ValidationDialog) and the data-validation
// alert (ValidationAlert), both from `editing`. SPEC-131: the drawing layer's
// surfaces, from the grid context's `drawings` — the Insert menu's image
// picker and shape gallery popover (SheetInsertSurfaces), the drawing context
// menu (SheetDrawingContextMenu) and the alt-text dialog
// (DrawingAltTextDialog). SPEC-133: the Name Box's names (the formula bar) and
// the Name Manager and its editor, from the grid context's `names`; the
// table surfaces (SheetTableSurfaces, the grid context's `tables`) — the
// design bar and the ODS notice sit BELOW the sheet, between it and the tabs,
// so they never move a cell under the pointer. SPEC-135: the link dialog
// (SheetLinkDialog, from `editing`'s links), which the cell menu's Link…
// opens too.
// Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { computed } from "vue";
import { asSheetGridContext } from "../../views/spreadsheet/useSheetGrid";
import { asWorkbookBook } from "../../views/spreadsheet/useWorkbookLoad";
import { asFunctionAssist } from "../../views/spreadsheet/useFunctionAssist";
import { namesOf } from "../../views/spreadsheet/useNames";
import { tablesOf } from "../../views/spreadsheet/useSheetTables";
import { sortFilterOf } from "../../views/spreadsheet/useSortFilter";
import { asWorkbookEditing } from "../../views/spreadsheet/useWorkbookEditing";
import DrawingAltTextDialog from "./DrawingAltTextDialog.vue";
import FormatCellsDialog from "./FormatCellsDialog.vue";
import FormulaBar from "./FormulaBar.vue";
import FunctionArgHint from "./FunctionArgHint.vue";
import FunctionAutocomplete from "./FunctionAutocomplete.vue";
import InsertFunctionDialog from "./InsertFunctionDialog.vue";
import PasteSpecialDialog from "./PasteSpecialDialog.vue";
import SheetCellContextMenu from "./SheetCellContextMenu.vue";
import SheetChartLayer from "./SheetChartLayer.vue";
import SheetChartSheet from "./SheetChartSheet.vue";
import SheetDrawingContextMenu from "./SheetDrawingContextMenu.vue";
import SheetGrid from "./SheetGrid.vue";
import SheetHeaderContextMenu from "./SheetHeaderContextMenu.vue";
import SheetInsertDeleteDialog from "./SheetInsertDeleteDialog.vue";
import SheetInsertSurfaces from "./SheetInsertSurfaces.vue";
import SheetLinkDialog from "./SheetLinkDialog.vue";
import SheetMergeWarningDialog from "./SheetMergeWarningDialog.vue";
import SheetMoveOverwriteDialog from "./SheetMoveOverwriteDialog.vue";
import SheetSidePanel from "./SheetSidePanel.vue";
import SheetSizeDialog from "./SheetSizeDialog.vue";
import SheetTableSurfaces from "./SheetTableSurfaces.vue";
import SheetTabColorPicker from "./SheetTabColorPicker.vue";
import SheetTabDeleteDialog from "./SheetTabDeleteDialog.vue";
import SheetTabMenu from "./SheetTabMenu.vue";
import SheetTabs from "./SheetTabs.vue";
import SheetUnhideDialog from "./SheetUnhideDialog.vue";
import SortFilterSurfaces from "./SortFilterSurfaces.vue";
import ValidationDialog from "./ValidationDialog.vue";
import ValidationAlert from "./ValidationAlert.vue";
import NameEditorDialog from "./NameEditorDialog.vue";
import NameManagerDialog from "./NameManagerDialog.vue";

const props = defineProps({
  /** useWorkbookLoad's return. */
  book: { type: Object, required: true },
  /** The SheetGridContext of the sheet on show. */
  grid: { type: Object, required: true },
  /** useWorkbookEditing's return (SPEC-121). */
  editing: { type: Object, required: true },
  /** useFunctionAssist's return (SPEC-127): the autocomplete, the argument
   * hint and the Insert Function dialog, and the formula bar's fx. */
  assist: { type: Object, default: null },
});

const assist = computed(() => (props.assist ? asFunctionAssist(props.assist) : null));

const book = computed(() => asWorkbookBook(props.book));

/** The cell menu's Insert…/Delete… (SPEC-124). */
function insertOrDelete(mode = "insert") {
  grid.value.structure?.insertOrDelete(mode === "delete" ? "delete" : "insert");
}
const grid = computed(() => asSheetGridContext(props.grid));
/** SPEC-128's sort and filter (null: not wired). */
const sortFilter = computed(() => sortFilterOf(props.grid));
/** SPEC-133: the Name Box, the Name Manager and its editor (null: not wired). */
const names = computed(() => namesOf(props.grid));
/** SPEC-133: the tables' design bar, dialogs and notice (null: not wired). */
const tables = computed(() => tablesOf(props.grid));

/** The cell menu's Cut, Copy, Paste and Paste special… (SPEC-125). */
function clipboard(action = "copy") {
  grid.value.clipboard?.menu(action === "cut" || action === "paste" || action === "paste-special" ? action : "copy");
}
const commands = computed(() => asWorkbookEditing(props.editing).commands);
const cellEditing = computed(() => asWorkbookEditing(props.editing).cell);
const validationDialog = computed(() => asWorkbookEditing(props.editing).validationDialog);
/** SPEC-130: data validation (absent when the editing wiring has none). */
const validation = computed(() => asWorkbookEditing(props.editing).validation ?? null);

/** SPEC-131: the drawing layer (null: not wired). */
const drawings = computed(() => grid.value.drawings ?? null);
/** SPEC-135: links — the link dialog, opened by the cell menu's Link… too (null: not wired). */
const links = computed(() => asWorkbookEditing(props.editing).links ?? null);
/** SPEC-132: the sheet charts — the toolbar, dialogs, gallery and chart sheets (null: not wired). */
const charts = computed(() => asWorkbookEditing(props.editing).charts ?? null);
function drawingAction(action = "front") {
  drawings.value?.runAction(action === "forward" || action === "backward" || action === "back" || action === "alt-text" || action === "delete" ? action : "front");
}

/** The validation alert's answer. */
function answerAlert(choice = "cancel") {
  validation.value?.answer(choice === "retry" || choice === "yes" || choice === "no" || choice === "ok" ? choice : "cancel");
}
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
      :editor="cellEditing"
      :editable="book.activeSheet.value?.kind === 'grid'"
      :can-insert-function="assist?.canInsert.value ?? false"
      :aria="assist?.editorAria.value ?? {}"
      :names="names ?? undefined"
      @insert-function="assist?.openDialog()"
    />
    <template v-if="assist">
      <FunctionAutocomplete
        v-if="assist.listOpen.value"
        :names="assist.matches.value"
        :highlighted="assist.highlighted.value ?? ''"
        :summary="assist.summary.value"
        :anchor-style="assist.anchorStyle.value"
        @choose="assist.choose"
      />
      <FunctionArgHint
        v-else-if="assist.hint.value"
        :parts="assist.hint.value.parts"
        :anchor-style="assist.anchorStyle.value"
      />
      <InsertFunctionDialog
        v-if="assist.dialogOpen.value"
        @insert="assist.insertFromDialog"
        @cancel="assist.closeDialog()"
      />
    </template>
    <div class="workbook-main">
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
        <SheetChartSheet
          v-else-if="book.activeSheet.value?.kind === 'chart'"
          :sheet="book.activeSheet.value"
          :theme="charts?.theme.value ?? null"
        />
        <SheetGrid
          v-else-if="book.activeSheet.value"
          :key="`${book.activeSheet.value.id}:${book.generation.value}`"
          :sheet="book.activeSheet.value"
          :context="grid"
          :show-formulas="grid.formulas.showFormulas.value"
          @cycle-sheet="book.cycleSheet"
        />
        <SheetChartLayer
          v-if="charts"
          :charts="charts"
        />
      </div>
      <SheetSidePanel :editing="editing" />
    </div>
    <SheetTableSurfaces
      v-if="tables"
      :tables="tables"
    />
    <SheetTabs
      :tabs="tabs"
      :active-index="book.activeIndex.value"
      :panel-id="panelId"
      :renaming-index="commands.renaming.value?.index ?? -1"
      :rename-draft="commands.renaming.value?.draft ?? ''"
      :rename-error="commands.renaming.value?.error ?? ''"
      :announcement="commands.announcement.value"
      @select="cellEditing.selectSheet"
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
    <template v-if="grid.formatting">
      <SheetCellContextMenu
        :open="grid.formatting.contextMenu.value !== null"
        :x="grid.formatting.contextMenu.value?.x ?? 0"
        :y="grid.formatting.contextMenu.value?.y ?? 0"
        @run="grid.formatting.run"
        @structure="insertOrDelete"
        @clipboard="clipboard"
        @link="links?.openDialog()"
        @close="grid.formatting.closeContextMenu()"
      />
      <FormatCellsDialog
        v-if="grid.formatting.dialog.open.value"
        :dialog="grid.formatting.dialog"
      />
    </template>
    <template v-if="grid.structure">
      <SheetHeaderContextMenu
        :open="grid.structure.headerMenu.value !== null"
        :x="grid.structure.headerMenu.value?.x ?? 0"
        :y="grid.structure.headerMenu.value?.y ?? 0"
        :items="grid.structure.headerItems.value"
        @run="grid.structure.runHeaderMenu"
        @close="grid.structure.closeHeaderMenu()"
      />
      <SheetInsertDeleteDialog
        v-if="grid.structure.insertDelete.value"
        :mode="grid.structure.insertDelete.value"
        @choose="grid.structure.chooseInsertDelete"
        @cancel="grid.structure.closeInsertDelete()"
      />
      <SheetSizeDialog
        v-if="grid.structure.sizeDialog.value"
        :axis="grid.structure.sizeDialog.value.axis"
        :initial="grid.structure.sizeDialog.value.initial"
        :mdw="grid.structure.sizeDialog.value.mdw"
        @confirm="grid.structure.applySize"
        @cancel="grid.structure.closeSize()"
      />
    </template>
    <SheetMergeWarningDialog
      v-if="grid.merges?.warning.value"
      @confirm="grid.merges.confirm()"
      @cancel="grid.merges.cancel()"
    />
    <PasteSpecialDialog
      v-if="grid.clipboard?.pasteSpecial.value"
      :capabilities="grid.clipboard.pasteSpecial.value.capabilities"
      @confirm="grid.clipboard.confirmPasteSpecial"
      @cancel="grid.clipboard.cancelPasteSpecial()"
    />
    <SheetMoveOverwriteDialog
      v-if="grid.rangeMove?.overwrite.value"
      @confirm="grid.rangeMove.confirmOverwrite()"
      @cancel="grid.rangeMove.cancelOverwrite()"
    />
    <SortFilterSurfaces
      v-if="sortFilter"
      :sort-filter="sortFilter"
    />
    <ValidationDialog
      v-if="validationDialog.open.value"
      :dialog="validationDialog"
    />
    <NameManagerDialog
      v-if="names?.managerOpen.value && !names.picking.value"
      :names="names"
    />
    <NameEditorDialog
      v-if="names?.editor.value"
      :names="names"
    />
    <SheetLinkDialog
      v-if="links?.dialog.value"
      :links="links"
    />
    <ValidationAlert
      v-if="validation?.alert.value"
      :alert-style="validation.alert.value.style"
      :title="validation.alert.value.title"
      :text="validation.alert.value.text"
      @answer="answerAlert"
    />
    <template v-if="drawings">
      <SheetInsertSurfaces
        :drawings="drawings"
        :charts="charts"
      />
      <SheetDrawingContextMenu
        :open="drawings.contextMenu.value !== null"
        :x="drawings.contextMenu.value?.x ?? 0"
        :y="drawings.contextMenu.value?.y ?? 0"
        :at-front="drawings.order.value.atFront"
        :at-back="drawings.order.value.atBack"
        @run="drawingAction"
        @close="drawings.closeContextMenu()"
      />
      <DrawingAltTextDialog
        v-if="drawings.altText.value"
        :initial="drawings.altText.value.initial"
        :name="drawings.altText.value.name"
        @confirm="drawings.confirmAltText"
        @cancel="drawings.cancelAltText()"
      />
    </template>
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

/* The sheet and SPEC-129's side panel, side by side. */
.workbook-main {
  flex: 1 1 auto;
  min-height: 0;
  display: flex;
}

.workbook-sheet {
  position: relative;
  flex: 1 1 auto;
  min-height: 0;
  min-width: 0;
}

.sheet-foreign {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--text-secondary);
}
</style>
