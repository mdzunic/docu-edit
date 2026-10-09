<script setup lang="ts">
// The workbook view at /sheets/:id (SPEC-116): a read-only wiring hub under
// the same rule as the other views (SpreadsheetView.size.spec.ts holds the
// line budget). Loading, the kind guard, the reader-version upgrade and the
// sheet on show live in useWorkbookLoad; the grid's viewport in
// useSheetViewport (inside SheetGrid); the surfaces in components/
// spreadsheet/*. This file only wires them under SPEC-058's app bar with
// SPEC-060's Close. Nothing here edits — no autosave session, no undo.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import SpreadsheetAppBar from "../components/spreadsheet/SpreadsheetAppBar.vue";
import SheetGrid from "../components/spreadsheet/SheetGrid.vue";
import SheetTabs from "../components/spreadsheet/SheetTabs.vue";
import { useTheme } from "../composables/useTheme";
import { useWorkbookLoad } from "./spreadsheet/useWorkbookLoad";

const route = useRoute();
const router = useRouter();

// No menu bar here — the shared theme composable keeps the view themed by
// the app tokens and following the OS live (SPEC-059).
useTheme();

const workbookId = computed(() => String(route.params.id ?? ""));
const book = useWorkbookLoad({ workbookId, router });
const panelId = "sheet-panel";
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
    <div
      v-if="book.workbook.value"
      class="workbook"
      data-testid="spreadsheet-workspace"
    >
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
          :key="book.activeSheet.value.id"
          :sheet="book.activeSheet.value"
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
    <div
      v-else
      class="workbook-missing"
      data-testid="spreadsheet-not-found"
      role="alert"
    >
      <p>This workbook doesn't exist in your library.</p>
    </div>
  </AppShell>
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

.sheet-foreign,
.workbook-missing {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--text-secondary);
}
</style>
