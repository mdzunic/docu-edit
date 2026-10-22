<script setup lang="ts">
// The workbook's app-bar cluster (SPEC-116 on SPEC-058's app bar): SPEC-121's
// undo/redo pair (`spreadsheet-undo`/`spreadsheet-redo`, tooltips "Undo
// Rename sheet (⌘Z)"), the format name ("Excel workbook (.xlsx)", …), the
// split Download — the file in its own format, or "Download a copy as …" the
// other (SPEC-126: both workbook formats for a csv/tsv workbook, and the
// current sheet as CSV or TSV for every one) — and SPEC-060's Close. AppShell's #app-bar-actions slot is where
// every editor puts its cluster; the title and the format badge are
// AppShell's own props. Runtime declarations per the espree constraint.
import UndoRedoControls from "../UndoRedoControls.vue";
import UiIconButton from "../ui/UiIconButton.vue";
import SpreadsheetDownloadButton from "./SpreadsheetDownloadButton.vue";

const props = defineProps({
  /** e.g. "Excel workbook (.xlsx)"; empty for a missing workbook. */
  formatLabel: { type: String, default: "" },
  canUndo: { type: Boolean, default: false },
  canRedo: { type: Boolean, default: false },
  /** The step the next undo/redo would report ("Rename sheet"). */
  undoLabel: { type: String, default: "" },
  redoLabel: { type: String, default: "" },
  /** The copy item's text, e.g. "Download a copy as OpenDocument
   * Spreadsheet (.ods)…". */
  copyLabel: { type: String, default: "" },
  /** Why the copy is unavailable; "" when it is available. */
  copyDisabledReason: { type: String, default: "" },
  /** SPEC-126: the .ods copy's label beside the .xlsx one (csv/tsv only). */
  copyOdsLabel: { type: String, default: "" },
});

const emit = defineEmits(["undo", "redo", "download", "download-copy", "download-copy-ods", "download-copy-csv", "download-copy-tsv", "export-pdf", "close"]);
</script>

<template>
  <div class="workbook-actions">
    <template v-if="props.formatLabel">
      <UndoRedoControls
        class="workbook-undo-redo"
        :can-undo="props.canUndo"
        :can-redo="props.canRedo"
        :undo-label="props.undoLabel"
        :redo-label="props.redoLabel"
        undo-testid="spreadsheet-undo"
        redo-testid="spreadsheet-redo"
        @undo="emit('undo')"
        @redo="emit('redo')"
      />
      <span
        class="workbook-format"
        data-testid="spreadsheet-format"
      >{{ props.formatLabel }}</span>
      <SpreadsheetDownloadButton
        :copy-label="props.copyLabel"
        :copy-disabled-reason="props.copyDisabledReason"
        :copy-ods-label="props.copyOdsLabel"
        @download="emit('download')"
        @download-copy="emit('download-copy')"
        @download-copy-ods="emit('download-copy-ods')"
        @download-copy-csv="emit('download-copy-csv')"
        @download-copy-tsv="emit('download-copy-tsv')"
        @export-pdf="emit('export-pdf')"
      />
    </template>
    <UiIconButton
      class="close-button"
      icon="close"
      label="Close document"
      data-testid="close-document"
      @click="emit('close')"
    />
  </div>
</template>

<style scoped>
.workbook-actions {
  display: flex;
  flex: 0 1 auto;
  min-width: 0;
  gap: var(--space-3);
  align-items: center;
}

/* The format name in the caption face, as the deck's facts are — chrome to
   glance at; it is what shrinks first on a crowded bar. */
.workbook-format {
  flex: 0 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
  white-space: nowrap;
}

/* On the brand bar the icon buttons — Close and the undo/redo pair — take
   the on-brand foreground and a translucent wash, the treatment every
   editor's Close gets. The pair's toolbar captions are dropped: the bar is a
   row of bare icons, each named by its tooltip. */
.workbook-actions .close-button,
.workbook-undo-redo :deep(.ui-icon-button) {
  color: var(--text-on-brand);
}

.workbook-undo-redo {
  --toolbar-caption-display: none;
}

.workbook-actions .close-button:not(:disabled):hover,
.workbook-undo-redo :deep(.ui-icon-button:not(:disabled):hover) {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.workbook-actions .close-button:not(:disabled):active,
.workbook-undo-redo :deep(.ui-icon-button:not(:disabled):active) {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}
</style>
