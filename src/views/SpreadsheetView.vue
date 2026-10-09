<script setup lang="ts">
// The workbook view at /sheets/:id (SPEC-116, SPEC-117): a read-only wiring
// hub under the same rule as the other views (SpreadsheetView.size.spec.ts
// holds the line budget). Loading, the kind guard, the reader-version upgrade
// and the sheet on show live in useWorkbookLoad; view state (selection, zoom,
// gridlines, scroll, active sheet — never a change to the file) in
// useSheetViewState; geometry in useSheetMetrics; zoom, selection and the
// menu bar's routing in useSheetZoom, useSheetSelection and
// useSpreadsheetMenu; the surfaces in components/spreadsheet/*. This file
// only wires them under SPEC-058's app bar with SPEC-060's Close.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import SpreadsheetAppBar from "../components/spreadsheet/SpreadsheetAppBar.vue";
import SpreadsheetMenuBar from "../components/spreadsheet/SpreadsheetMenuBar.vue";
import SpreadsheetStatusBar from "../components/spreadsheet/SpreadsheetStatusBar.vue";
import WorkbookBody from "../components/spreadsheet/WorkbookBody.vue";
import { useTheme } from "../composables/useTheme";
import { useSheetMetrics } from "./spreadsheet/useSheetMetrics";
import { useSheetSelection } from "./spreadsheet/useSheetSelection";
import { useSheetViewState } from "./spreadsheet/useSheetViewState";
import { useSheetZoom } from "./spreadsheet/useSheetZoom";
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
const selection = useSheetSelection({ sheet: () => book.activeSheet.value, sheetMetrics, viewState });
const menu = useSpreadsheetMenu({ book, viewState, zoom });
const grid = computed(() => ({ model: book.workbook.value, sheetMetrics, selection, zoom, viewState }));
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
