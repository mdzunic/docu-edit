<script setup lang="ts">
// The workbook view at /sheets/:id (SPEC-116, SPEC-117, SPEC-121): a wiring
// hub (SpreadsheetView.size.spec.ts holds its line budget). Loading, the kind
// guard, the reader upgrade and the sheet on show: useWorkbookLoad; view state
// (never a change to the file): useSheetViewState; geometry, zoom, selection,
// menus, the locale (SPEC-118), formula display (SPEC-119), recalculation
// (SPEC-120), the sheet commands, undo and downloads (SPEC-121) and cell entry
// (SPEC-122) — useWorkbookEditing: the spreadsheet/ composables below; the
// surfaces: components/spreadsheet/*. This file only wires them.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import SpreadsheetAppBar from "../components/spreadsheet/SpreadsheetAppBar.vue";
import SpreadsheetMenuBar from "../components/spreadsheet/SpreadsheetMenuBar.vue";
import SpreadsheetStatusBar from "../components/spreadsheet/SpreadsheetStatusBar.vue";
import WorkbookBody from "../components/spreadsheet/WorkbookBody.vue";
import WorkbookNotFound from "../components/spreadsheet/WorkbookNotFound.vue";
import { useTheme } from "../composables/useTheme";
import { useFormulaDisplay } from "./spreadsheet/useFormulaDisplay";
import { useRecalc } from "./spreadsheet/useRecalc";
import { useSheetMetrics } from "./spreadsheet/useSheetMetrics";
import { useSheetSelection } from "./spreadsheet/useSheetSelection";
import { useSheetViewState } from "./spreadsheet/useSheetViewState";
import { useSheetZoom } from "./spreadsheet/useSheetZoom";
import { useSpreadsheetLocale } from "./spreadsheet/useSpreadsheetLocale";
import { useSpreadsheetMenu } from "./spreadsheet/useSpreadsheetMenu";
import { useWorkbookEditing } from "./spreadsheet/useWorkbookEditing";
import { useWorkbookLoad } from "./spreadsheet/useWorkbookLoad";

const route = useRoute();
const router = useRouter();

// The shared theme composable keeps the view themed by the app tokens and
// following the OS live (SPEC-059).
useTheme();

const workbookId = computed(() => String(route.params.id ?? ""));
const book = useWorkbookLoad({ workbookId, router });
const viewState = useSheetViewState({ workbook: book.workbook, activeIndex: book.activeIndex });
const sheetMetrics = useSheetMetrics(() => book.activeSheet.value);
const zoom = useSheetZoom({ viewState });
const sheetLocale = useSpreadsheetLocale();
const selection = useSheetSelection({ sheet: () => book.activeSheet.value, sheetMetrics, viewState, model: () => book.workbook.value, locale: () => sheetLocale.locale.value });
const formulas = useFormulaDisplay({ sheet: () => book.activeSheet.value, model: () => book.workbook.value, viewState, selection });
const recalc = useRecalc({ book, locale: () => sheetLocale.locale.value });
const editing = useWorkbookEditing({ book, viewState, recalc, selection, sheetMetrics, locale: () => sheetLocale.locale.value });
const menu = useSpreadsheetMenu({ book, viewState, zoom, formulas, recalc, editing });
const grid = computed(() => ({ model: book.workbook.value, sheetMetrics, selection, zoom, viewState, locale: sheetLocale.locale.value, formulas, recalc, cellEditing: editing.cell }));
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
      <SpreadsheetMenuBar
        :state="menu.state.value"
        @command="menu.run"
      />
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
        v-bind="{ ...selection.statusTexts.value, ...recalc.status.value, mode: editing.cell.modeLabel.value }"
        :zoom-percent="zoom.percent.value"
        :zoom-disabled="book.activeSheet.value?.kind !== 'grid'"
        @zoom-in="zoom.zoomIn"
        @zoom-out="zoom.zoomOut"
        @apply-zoom="zoom.applyInput"
        @set-zoom="zoom.set"
      />
    </template>
  </AppShell>
</template>
