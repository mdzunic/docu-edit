<script setup lang="ts">
// The workbook view at /sheets/:id, a wiring hub (SpreadsheetView.size.spec.ts holds its budget): each composable owns its feature —
// loading, view state, selection, editing, formatting, rows and columns, merges, clipboard, fill and move; surfaces: components/.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import SpreadsheetAppBar from "../components/spreadsheet/SpreadsheetAppBar.vue";
import SpreadsheetStatusBar from "../components/spreadsheet/SpreadsheetStatusBar.vue";
import WorkbookBody from "../components/spreadsheet/WorkbookBody.vue";
import WorkbookCommandBar from "../components/spreadsheet/WorkbookCommandBar.vue";
import WorkbookNotFound from "../components/spreadsheet/WorkbookNotFound.vue";
import { useTheme } from "../composables/useTheme";
import { useFillHandle } from "./spreadsheet/useFillHandle";
import { useFormatPainter } from "./spreadsheet/useFormatPainter";
import { useFormatting } from "./spreadsheet/useFormatting";
import { useFormulaDisplay } from "./spreadsheet/useFormulaDisplay";
import { useHeaderResize } from "./spreadsheet/useHeaderResize";
import { useMerges } from "./spreadsheet/useMerges";
import { useRecalc } from "./spreadsheet/useRecalc";
import { useRangeMove } from "./spreadsheet/useRangeMove";
import { useSheetClipboard } from "./spreadsheet/useSheetClipboard";
import { useSheetMetrics } from "./spreadsheet/useSheetMetrics";
import { useSheetSelection } from "./spreadsheet/useSheetSelection";
import { useSheetViewState } from "./spreadsheet/useSheetViewState";
import { useSheetZoom } from "./spreadsheet/useSheetZoom";
import { useSpreadsheetLocale } from "./spreadsheet/useSpreadsheetLocale";
import { useSpreadsheetMenu } from "./spreadsheet/useSpreadsheetMenu";
import { useStructure } from "./spreadsheet/useStructure";
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
const formatting = useFormatting({ book, viewState, selection, editing, locale: () => sheetLocale.locale.value });
const structure = useStructure({ book, viewState, selection, sheetMetrics, editing });
const headerResize = useHeaderResize({ book, structure, selection, sheetMetrics, zoom });
const merges = useMerges({ book, viewState, selection, sheetMetrics, editing });
const painter = useFormatPainter({ book, selection, formatting, merges });
const clipboard = useSheetClipboard({ book, viewState, selection, editing, locale: () => sheetLocale.locale.value });
const fillHandle = useFillHandle({ book, viewState, selection, editing, locale: () => sheetLocale.locale.value });
const rangeMove = useRangeMove({ book, viewState, selection, editing });
const menu = useSpreadsheetMenu({ book, viewState, zoom, formulas, recalc, editing, formatting, structure, clipboard, fillHandle });
const grid = computed(() => ({ model: book.workbook.value, sheetMetrics, selection, zoom, viewState, locale: sheetLocale.locale.value, formulas, recalc, cellEditing: editing.cell, formatting, painter, structure, headerResize, merges, clipboard, fillHandle, rangeMove }));
</script>

<template>
  <AppShell
    :document-title="book.title.value"
    :format-badge="book.badge.value"
  >
    <template #app-bar-actions>
      <SpreadsheetAppBar
        :format-label="book.formatLabel.value"
        v-bind="editing.appBar.value"
        @undo="editing.undo.undo()"
        @redo="editing.undo.redo()"
        @download="editing.download.download()"
        @download-copy="editing.download.downloadCopy()"
        @close="book.close()"
      />
    </template>
    <template
      v-if="book.workbook.value"
      #menubar
    >
      <WorkbookCommandBar v-bind="{ menu, formatting, painter, merges }" />
    </template>
    <WorkbookBody
      v-if="book.workbook.value"
      :book="book"
      :grid="grid"
      :editing="editing"
    />
    <WorkbookNotFound v-else />
    <template
      v-if="book.workbook.value"
      #status-bar
    >
      <SpreadsheetStatusBar
        v-bind="{ ...selection.statusTexts.value, ...recalc.status.value, mode: editing.cell.modeLabel.value, zoomPercent: zoom.percent.value, zoomDisabled: book.activeSheet.value?.kind !== 'grid' }"
        @zoom-in="zoom.zoomIn"
        @zoom-out="zoom.zoomOut"
        @apply-zoom="zoom.applyInput"
        @set-zoom="zoom.set"
      />
    </template>
  </AppShell>
</template>
