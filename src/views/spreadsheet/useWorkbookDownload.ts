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
import { computed, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { spreadsheetFormatLabel } from "../../lib/format";
import { downloadSpreadsheet } from "../../lib/spreadsheet/download";
import type { SpreadsheetFormat, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { useSpreadsheetLocale } from "./useSpreadsheetLocale";

/** "Download a copy as OpenDocument Spreadsheet (.ods)…" / "…as Excel workbook (.xlsx)…". */
export function downloadCopyLabel(target: SpreadsheetFormat): string {
  return `Download a copy as ${spreadsheetFormatLabel(target)}…`;
}

export interface WorkbookDownloadDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeIndex: Ref<number> };
}

export function useWorkbookDownload(deps: WorkbookDownloadDeps) {
  const { addToast } = useToasts();
  const { locale } = useSpreadsheetLocale();

  const otherFormat = computed((): SpreadsheetFormat => (deps.book.workbook.value?.format === "ods" ? "xlsx" : "ods"));
  const copyLabel = computed(() => downloadCopyLabel(otherFormat.value));
  const copyDisabledReason = computed(() => "");

  /** The model as the download writes it: the sheet on show active. */
  function viewed(model: SpreadsheetModel): SpreadsheetModel {
    return model.activeSheetIndex === deps.book.activeIndex.value ? model : { ...model, activeSheetIndex: deps.book.activeIndex.value };
  }

  function save(target?: SpreadsheetFormat): void {
    const model = deps.book.workbook.value;
    if (!model) {
      return;
    }
    void downloadSpreadsheet(viewed(model), target, locale.value).then((notices) => {
      for (const notice of notices) {
        addToast(notice);
      }
    });
  }

  return {
    copyLabel,
    copyDisabledReason,
    download: () => save(),
    downloadCopy: () => {
      if (copyDisabledReason.value === "") {
        save(otherFormat.value);
      }
    },
  };
}

export type WorkbookDownload = ReturnType<typeof useWorkbookDownload>;
