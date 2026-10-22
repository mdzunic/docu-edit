<script setup lang="ts">
// The workbook view's app-bar wiring: SpreadsheetAppBar bound to the view's
// load and editing composables — undo/redo, Download and its menu's copies
// (SPEC-121's other format; SPEC-126's .ods sibling of a csv/tsv workbook's
// .xlsx copy and the current sheet as CSV or TSV) and Close. Extracted so
// SpreadsheetView keeps its line budget (SpreadsheetView.size.spec.ts), the
// move WorkbookCommandBar made. SPEC-135: `sheet-comments-toggle` before it
// toggles the side panel's comments pane (useWorkbookEditing's `comments`).
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asWorkbookBook } from "../../views/spreadsheet/useWorkbookLoad";
import { asWorkbookEditing } from "../../views/spreadsheet/useWorkbookEditing";
import UiIconButton from "../ui/UiIconButton.vue";
import SpreadsheetAppBar from "./SpreadsheetAppBar.vue";

const props = defineProps({
  /** useWorkbookLoad's return. */
  book: { type: Object, required: true },
  /** useWorkbookEditing's return. */
  editing: { type: Object, required: true },
});

const book = computed(() => asWorkbookBook(props.book));
const editing = computed(() => asWorkbookEditing(props.editing));
</script>

<template>
  <UiIconButton
    v-if="editing.comments && book.formatLabel.value"
    icon="comments"
    label="Comments"
    data-testid="sheet-comments-toggle"
    :aria-pressed="editing.comments.isOpen.value ? 'true' : 'false'"
    @click="editing.comments.toggle()"
  />
  <SpreadsheetAppBar
    :format-label="book.formatLabel.value"
    v-bind="editing.appBar.value"
    @undo="editing.undo.undo()"
    @redo="editing.undo.redo()"
    @download="editing.download.download()"
    @download-copy="editing.download.downloadCopy()"
    @download-copy-ods="editing.download.downloadCopyOds()"
    @download-copy-csv="editing.download.downloadSheetCopy('csv')"
    @download-copy-tsv="editing.download.downloadSheetCopy('tsv')"
    @close="book.close()"
  />
</template>
