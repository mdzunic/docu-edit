<script setup lang="ts">
// The workbook view at /sheets/:id, a wiring hub (SpreadsheetView.size.spec.ts holds its budget): each composable owns its feature — loading,
// view state, saving (autosave, recovery, edit lock), selection, editing, the grid's tools (useSheetTools), find, Go To, history; surfaces: components/.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import CsvImportBar from "../components/spreadsheet/CsvImportBar.vue";
import CsvImportDialog from "../components/spreadsheet/CsvImportDialog.vue";
import GoToDialog from "../components/spreadsheet/GoToDialog.vue";
import GoToSpecialDialog from "../components/spreadsheet/GoToSpecialDialog.vue";
import WorkbookAppBar from "../components/spreadsheet/WorkbookAppBar.vue";
import WorkbookBody from "../components/spreadsheet/WorkbookBody.vue";
import WorkbookCommandBar from "../components/spreadsheet/WorkbookCommandBar.vue";
import WorkbookFindBar from "../components/spreadsheet/WorkbookFindBar.vue";
import WorkbookNotFound from "../components/spreadsheet/WorkbookNotFound.vue";
import WorkbookStatusBar from "../components/spreadsheet/WorkbookStatusBar.vue";
import { useTheme } from "../composables/useTheme";
import { useFormulaDisplay } from "./spreadsheet/useFormulaDisplay";
import { useRecalc } from "./spreadsheet/useRecalc";
import { useSheetFind } from "./spreadsheet/useSheetFind";
import { useSheetGoTo } from "./spreadsheet/useSheetGoTo";
import { useSheetMetrics } from "./spreadsheet/useSheetMetrics";
import { useSheetSelection } from "./spreadsheet/useSheetSelection";
import { useSheetTools } from "./spreadsheet/useSheetTools";
import { useSheetViewState } from "./spreadsheet/useSheetViewState";
import { useSheetZoom } from "./spreadsheet/useSheetZoom";
import { useSpreadsheetLocale } from "./spreadsheet/useSpreadsheetLocale";
import { useWorkbookEditing } from "./spreadsheet/useWorkbookEditing";
import { useWorkbookHistory } from "./spreadsheet/useWorkbookHistory";
import { useWorkbookLoad } from "./spreadsheet/useWorkbookLoad";
import { useWorkbookSaving } from "./spreadsheet/useWorkbookSaving";

const route = useRoute();
useTheme(); // the app tokens' theme, following the OS (SPEC-059)

const book = useWorkbookLoad({ workbookId: computed(() => String(route.params.id ?? "")), router: useRouter() });
const viewState = useSheetViewState({ workbook: book.workbook, activeIndex: book.activeIndex, writable: () => saving.owns.value && !saving.previewing.value });
const saving = useWorkbookSaving({ model: book.workbook, workbookId: book.workbookId, getSheetView: viewState.sheetView, applySheetView: viewState.applySheetView, resetUndo: () => editing.undo.reset(), download: () => editing.download.download(), show: book.replace, remount: book.remount, storageIncomplete: book.storageIncomplete });
const sheetMetrics = useSheetMetrics(() => book.activeSheet.value);
const zoom = useSheetZoom({ viewState });
const sheetLocale = useSpreadsheetLocale();
const selection = useSheetSelection({ sheet: () => book.activeSheet.value, sheetMetrics, viewState, model: () => book.workbook.value, locale: () => sheetLocale.locale.value });
const formulas = useFormulaDisplay({ sheet: () => book.activeSheet.value, model: () => book.workbook.value, viewState, selection });
const recalc = useRecalc({ book, locale: () => sheetLocale.locale.value });
const editing = useWorkbookEditing({ book, viewState, recalc, selection, sheetMetrics, locale: () => sheetLocale.locale.value, session: saving });
const { formatting, structure, headerResize, merges, painter, clipboard, fillHandle, rangeMove, sortFilter, names, tables, menu, printing } = useSheetTools({ book, viewState, selection, sheetMetrics, zoom, formulas, recalc, editing, locale: () => sheetLocale.locale.value });
const find = useSheetFind({ book, viewState, selection, editing, locale: () => sheetLocale.locale.value });
const goTo = useSheetGoTo({ book, viewState, selection, editing, names });
const history = useWorkbookHistory({ saving, model: book.workbook, show: book.replace, syncView: viewState.syncInto, resetUndo: () => editing.undo.reset(), sidePanel: editing.sidePanel });
const grid = computed(() => ({ model: book.workbook.value, sheetMetrics, selection, zoom, viewState, locale: sheetLocale.locale.value, formulas, recalc, cellEditing: editing.gridEditor.value, reveal: editing.cell.revealTick, conditionalFormats: editing.conditionalFormats, assist: editing.assist, formatting, painter, structure, headerResize, merges, clipboard, fillHandle, rangeMove, sortFilter, validation: editing.validation, drawings: editing.drawings, notes: editing.notes, links: editing.links, names, tables, find, printing }));
</script>

<template>
  <AppShell
    :document-title="book.title.value"
    :format-badge="book.badge.value"
  >
    <template #app-bar-actions>
      <WorkbookAppBar v-bind="{ book, editing, printing, saving, history }" />
    </template>
    <template
      v-if="book.workbook.value"
      #menubar
    >
      <WorkbookCommandBar v-bind="{ menu, formatting, painter, merges }" />
      <CsvImportBar :csv-import="editing.csvImport" />
    </template>
    <WorkbookFindBar v-bind="{ find, readOnly: editing.readOnly.value }" />
    <WorkbookBody
      v-if="book.workbook.value"
      v-bind="{ book, grid, editing, assist: editing.assist, saving, history }"
    />
    <WorkbookNotFound v-else />
    <CsvImportDialog :csv-import="editing.csvImport" />
    <GoToDialog :go-to="goTo" />
    <GoToSpecialDialog :go-to="goTo" />
    <template
      v-if="book.workbook.value"
      #status-bar
    >
      <WorkbookStatusBar v-bind="{ book, selection, recalc, editing, zoom, sortFilter }" />
    </template>
  </AppShell>
</template>
