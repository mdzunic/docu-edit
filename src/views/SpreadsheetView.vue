<script setup lang="ts">
// The workbook view at /sheets/:id, a wiring hub (SpreadsheetView.size.spec.ts holds its budget): each composable owns its feature —
// loading, view state, selection, editing (drawings, CSV import included), the grid's tools (useSheetTools), find, Go To; surfaces: components/.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import FindBar from "../components/FindBar.vue";
import CsvImportBar from "../components/spreadsheet/CsvImportBar.vue";
import CsvImportDialog from "../components/spreadsheet/CsvImportDialog.vue";
import GoToDialog from "../components/spreadsheet/GoToDialog.vue";
import GoToSpecialDialog from "../components/spreadsheet/GoToSpecialDialog.vue";
import SheetFindOptions from "../components/spreadsheet/SheetFindOptions.vue";
import SheetFindResults from "../components/spreadsheet/SheetFindResults.vue";
import WorkbookAppBar from "../components/spreadsheet/WorkbookAppBar.vue";
import WorkbookBody from "../components/spreadsheet/WorkbookBody.vue";
import WorkbookCommandBar from "../components/spreadsheet/WorkbookCommandBar.vue";
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
import { useWorkbookLoad } from "./spreadsheet/useWorkbookLoad";

const route = useRoute();
useTheme(); // the app tokens' theme, following the OS (SPEC-059)

const book = useWorkbookLoad({ workbookId: computed(() => String(route.params.id ?? "")), router: useRouter() });
const viewState = useSheetViewState({ workbook: book.workbook, activeIndex: book.activeIndex });
const sheetMetrics = useSheetMetrics(() => book.activeSheet.value);
const zoom = useSheetZoom({ viewState });
const sheetLocale = useSpreadsheetLocale();
const selection = useSheetSelection({ sheet: () => book.activeSheet.value, sheetMetrics, viewState, model: () => book.workbook.value, locale: () => sheetLocale.locale.value });
const formulas = useFormulaDisplay({ sheet: () => book.activeSheet.value, model: () => book.workbook.value, viewState, selection });
const recalc = useRecalc({ book, locale: () => sheetLocale.locale.value });
const editing = useWorkbookEditing({ book, viewState, recalc, selection, sheetMetrics, locale: () => sheetLocale.locale.value });
const { formatting, structure, headerResize, merges, painter, clipboard, fillHandle, rangeMove, sortFilter, names, tables, menu } = useSheetTools({ book, viewState, selection, sheetMetrics, zoom, formulas, recalc, editing, locale: () => sheetLocale.locale.value });
const find = useSheetFind({ book, viewState, selection, editing, locale: () => sheetLocale.locale.value });
const goTo = useSheetGoTo({ book, viewState, selection, editing, names });
const grid = computed(() => ({ model: book.workbook.value, sheetMetrics, selection, zoom, viewState, locale: sheetLocale.locale.value, formulas, recalc, cellEditing: editing.cell, conditionalFormats: editing.conditionalFormats, assist: editing.assist, formatting, painter, structure, headerResize, merges, clipboard, fillHandle, rangeMove, sortFilter, validation: editing.validation, drawings: editing.drawings, notes: editing.notes, links: editing.links, names, tables, find }));
</script>

<template>
  <AppShell
    :document-title="book.title.value"
    :format-badge="book.badge.value"
  >
    <template #app-bar-actions>
      <WorkbookAppBar v-bind="{ book, editing }" />
    </template>
    <template
      v-if="book.workbook.value"
      #menubar
    >
      <WorkbookCommandBar v-bind="{ menu, formatting, painter, merges }" />
      <CsvImportBar :csv-import="editing.csvImport" />
    </template>
    <SheetFindResults :find="find">
      <FindBar
        v-bind="find.bar.value"
        v-on="find.barEvents"
      >
        <template #options>
          <SheetFindOptions :find="find" />
        </template>
      </FindBar>
    </SheetFindResults>
    <WorkbookBody
      v-if="book.workbook.value"
      v-bind="{ book, grid, editing, assist: editing.assist }"
    />
    <WorkbookNotFound v-else :loading="book.loading.value" />
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
