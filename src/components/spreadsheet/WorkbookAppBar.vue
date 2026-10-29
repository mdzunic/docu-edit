<script setup lang="ts">
// The workbook view's app-bar wiring: SpreadsheetAppBar bound to the view's
// load and editing composables — undo/redo, Download and its menu's copies
// (SPEC-121's other format; SPEC-126's .ods sibling of a csv/tsv workbook's
// .xlsx copy and the current sheet as CSV or TSV) and Close. Extracted so
// SpreadsheetView keeps its line budget (SpreadsheetView.size.spec.ts), the
// move WorkbookCommandBar made. SPEC-135: `sheet-comments-toggle` before it
// toggles the side panel's comments pane (useWorkbookEditing's `comments`).
// SPEC-138: `sheet-history-toggle` toggles its version history pane
// (useWorkbookHistory), and the bar shows useWorkbookSaving's save state.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asWorkbookBook } from "../../views/spreadsheet/useWorkbookLoad";
import { asWorkbookEditing } from "../../views/spreadsheet/useWorkbookEditing";
import { asWorkbookHistory } from "../../views/spreadsheet/useWorkbookHistory";
import { asWorkbookSaving } from "../../views/spreadsheet/useWorkbookSaving";
import UiIconButton from "../ui/UiIconButton.vue";
import SpreadsheetAppBar from "./SpreadsheetAppBar.vue";
import { asSheetPrinting } from "../../views/spreadsheet/useSheetPrint";

const props = defineProps({
  /** useWorkbookLoad's return. */
  book: { type: Object, required: true },
  /** useWorkbookEditing's return. */
  editing: { type: Object, required: true },
  /** SPEC-136: the view's SheetPrinting (the download menu's "PDF (.pdf)…"). */
  printing: { type: Object, default: null },
  /** SPEC-138: useWorkbookSaving's return (the save state). */
  saving: { type: Object, default: null },
  /** SPEC-138: useWorkbookHistory's return (the history toggle). */
  history: { type: Object, default: null },
});

const book = computed(() => asWorkbookBook(props.book));
const editing = computed(() => asWorkbookEditing(props.editing));
const saving = computed(() => (props.saving ? asWorkbookSaving(props.saving) : null));
const history = computed(() => (props.history ? asWorkbookHistory(props.history) : null));
/** The badge's state and words — none while another tab holds the workbook. */
const saveBadge = computed(() =>
  saving.value?.showSaveState.value ? { saveState: saving.value.saveState.value, saveStateText: saving.value.saveStateText.value } : {},
);

function exportPdf() {
  if (props.printing) {
    void asSheetPrinting(props.printing).pdf.exportPdf();
  }
}
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
  <UiIconButton
    v-if="history && book.formatLabel.value"
    icon="history"
    label="Version history"
    data-testid="sheet-history-toggle"
    :aria-pressed="history.isOpen.value ? 'true' : 'false'"
    @click="history.toggle()"
  />
  <SpreadsheetAppBar
    :format-label="book.formatLabel.value"
    v-bind="{ ...editing.appBar.value, ...saveBadge }"
    @undo="editing.undo.undo()"
    @redo="editing.undo.redo()"
    @download="editing.download.download()"
    @download-copy="editing.download.downloadCopy()"
    @download-copy-ods="editing.download.downloadCopyOds()"
    @download-copy-csv="editing.download.downloadSheetCopy('csv')"
    @download-copy-tsv="editing.download.downloadSheetCopy('tsv')"
    @export-pdf="exportPdf()"
    @close="book.close()"
  />
</template>
