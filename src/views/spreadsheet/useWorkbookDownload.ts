// The workbook's downloads (SPEC-121 "Cross-format copy", AC-20): the app
// bar's split "Download ▾" and File ▸ Download / "Download a copy as …".
// Download writes the workbook in its own format (lib/spreadsheet/download.ts
// picks the writer's path); the copy converts through stripRetention into
// the other format and shows the notices it returns as toasts ("Not
// included in the .ods copy: charts, comments"). Both write the sheet on
// show as the active one, and an .ods in the user's number locale (SPEC-122:
// a cell's informative text). SPEC-122's from-scratch writers serialise
// every sheet's cells and styles, so the copy is offered for every workbook
// (`copyDisabledReason` stays "" — the prop plumbing keeps a reason
// possible for a later spec).
//
// SPEC-126: a csv/tsv workbook's copies are both workbook formats — the
// .xlsx one under `copyLabel` and an .ods sibling under `copyOdsLabel` —
// each through stripRetention's from-scratch path; and every workbook gains
// "Download a copy as CSV/TSV (current sheet)" (downloadSheetCopy): the
// sheet on show in the user's locale, its losses as a toast.
import { computed, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { spreadsheetFormatLabel } from "../../lib/format";
import { downloadSpreadsheet, downloadSpreadsheetSheetCopy } from "../../lib/spreadsheet/download";
import { isCsvFormat, type SpreadsheetFormat, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import { useSpreadsheetLocale } from "./useSpreadsheetLocale";

/** "Download a copy as OpenDocument Spreadsheet (.ods)…" / "…as Excel workbook (.xlsx)…". */
export function downloadCopyLabel(target: SpreadsheetFormat): string {
  return `Download a copy as ${spreadsheetFormatLabel(target)}…`;
}

export interface WorkbookDownloadDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeIndex: Ref<number> };
  /** SPEC-138: a download is SPEC-010's explicit save — the saving
   * session's commit point (commit, then SPEC-040's snapshot). */
  onDownload?: () => void;
}

export function useWorkbookDownload(deps: WorkbookDownloadDeps) {
  const { addToast } = useToasts();
  const { locale } = useSpreadsheetLocale();

  const format = computed(() => deps.book.workbook.value?.format ?? "xlsx");
  /** The copy the main item makes: the other package format, or .xlsx for
   * a csv/tsv workbook. */
  const otherFormat = computed((): SpreadsheetFormat => (format.value === "xlsx" ? "ods" : "xlsx"));
  const copyLabel = computed(() => downloadCopyLabel(otherFormat.value));
  const copyDisabledReason = computed(() => "");
  /** SPEC-126: the .ods sibling of a csv/tsv workbook's .xlsx copy. */
  const copyOdsLabel = computed(() => (isCsvFormat(format.value) ? downloadCopyLabel("ods") : ""));

  /** The model as the download writes it: the sheet on show active. */
  function viewed(model: SpreadsheetModel): SpreadsheetModel {
    return model.activeSheetIndex === deps.book.activeIndex.value ? model : { ...model, activeSheetIndex: deps.book.activeIndex.value };
  }

  function toast(notices: string[]): void {
    for (const notice of notices) {
      addToast(notice);
    }
  }

  function save(target?: SpreadsheetFormat): void {
    const model = deps.book.workbook.value;
    if (!model) {
      return;
    }
    void downloadSpreadsheet(viewed(model), target, locale.value).then(toast);
    deps.onDownload?.();
  }

  /** "Download a copy as CSV/TSV (current sheet)": the sheet on show. */
  function downloadSheetCopy(target: "csv" | "tsv"): void {
    const model = deps.book.workbook.value;
    if (model) {
      void downloadSpreadsheetSheetCopy(viewed(model), target, locale.value).then(toast);
      deps.onDownload?.();
    }
  }

  return {
    copyLabel,
    copyDisabledReason,
    copyOdsLabel,
    download: () => save(),
    downloadCopy: () => {
      if (copyDisabledReason.value === "") {
        save(otherFormat.value);
      }
    },
    downloadCopyOds: () => save("ods"),
    downloadSheetCopy,
  };
}

export type WorkbookDownload = ReturnType<typeof useWorkbookDownload>;
