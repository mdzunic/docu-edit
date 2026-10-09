<script setup lang="ts">
// The workbook view at /sheets/:id (SPEC-116, SPEC-117): a read-only wiring
// hub (SpreadsheetView.size.spec.ts holds its line budget). Loading, the kind
// guard, the reader upgrade and the sheet on show: useWorkbookLoad; view state
// (never a change to the file): useSheetViewState; geometry, zoom, selection,
// menu routing, the number locale (SPEC-118) and formula display (SPEC-119):
// the spreadsheet/ composables imported below; the surfaces:
// components/spreadsheet/*. This file only wires them.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import SpreadsheetAppBar from "../components/spreadsheet/SpreadsheetAppBar.vue";
import SpreadsheetMenuBar from "../components/spreadsheet/SpreadsheetMenuBar.vue";
import SpreadsheetStatusBar from "../components/spreadsheet/SpreadsheetStatusBar.vue";
import WorkbookBody from "../components/spreadsheet/WorkbookBody.vue";
import { useTheme } from "../composables/useTheme";
import { useFormulaDisplay } from "./spreadsheet/useFormulaDisplay";
import { useSheetMetrics } from "./spreadsheet/useSheetMetrics";
import { useSheetSelection } from "./spreadsheet/useSheetSelection";
import { useSheetViewState } from "./spreadsheet/useSheetViewState";
import { useSheetZoom } from "./spreadsheet/useSheetZoom";
import { useSpreadsheetLocale } from "./spreadsheet/useSpreadsheetLocale";
import { useSpreadsheetMenu } from "./spreadsheet/useSpreadsheetMenu";
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
const menu = useSpreadsheetMenu({ book, viewState, zoom, formulas });
const grid = computed(() => ({ model: book.workbook.value, sheetMetrics, selection, zoom, viewState, locale: sheetLocale.locale.value, formulas }));
</script>

<template>
  <AppShell
    :document-title="book.title.value"
    :format-badge="book.badge.value"
  >
    <template #app-bar-actions>
      <SpreadsheetAppBar
        :format-label="book.formatLabel.value"
        @download="book.download()"
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
    />
    <div
      v-else
      class="workbook-missing"
      data-testid="spreadsheet-not-found"
      role="alert"
    >
      <p>This workbook doesn't exist in your library.</p>
    </div>
    <template
      v-if="book.workbook.value"
      #status-bar
    >
      <SpreadsheetStatusBar
        :average="selection.statusTexts.value.average"
        :count="selection.statusTexts.value.count"
        :sum="selection.statusTexts.value.sum"
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

<style scoped>
.workbook-missing {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--text-secondary);
}
</style>
