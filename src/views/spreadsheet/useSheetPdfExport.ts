// SPEC-136's PDF export for the workbook view ("Download PDF" in the print
// preview, the download menu's "PDF (.pdf)…"): a real PDF through the core
// writer (SPEC-104's core/pdfBuild.ts, core/pdfExtract.ts and the shared
// face and image loaders of lib/pdfAssets.ts), never a screenshot.
//
// - Pages are mounted, extracted and unmounted off-screen at 1:1 — the
//   `mode="sheet"` SheetPrintPages, fixed at left −20000px, aria-hidden and
//   inert; never the preview's fit-scaled pages, because extractPageContent
//   multiplies computed font sizes by the page's on-screen scale — in a
//   rolling window of EXPORT_WINDOW_PAGES, yieldToMainThread() between pages.
// - Each page: extractPageContent(page, widthPt, heightPt, { cellSelector:
//   ".sheet-print-cell", clipSelector: ".sheet-print-cell", imageSelector,
//   canvasSelector }); one level-1 heading per sheet (its name) on its first
//   page; `/Title` the workbook's title; saved as pdfFileName(title,
//   DEFAULT_SPREADSHEET_TITLE).
// - `progress` reads "Page {done} of {total}"; abort() stops at the next page
//   boundary with no download; an export over MAX_SHEET_PDF_PAGES is refused.
//   Nothing here writes the model.
import { nextTick, ref, shallowRef } from "vue";
import { buildPdf, PDF_MIME_TYPE, pdfFileName, type PdfPageInput } from "../../lib/core/pdfBuild";
import { extractPageContent } from "../../lib/core/pdfExtract";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { loadPdfFonts, loadPdfImages } from "../../lib/pdfAssets";
import { DEFAULT_SPREADSHEET_TITLE } from "../../lib/spreadsheet/model";
import { EXPORT_WINDOW_PAGES, MAX_SHEET_PDF_PAGES, PDF_LIMIT_MESSAGE } from "../../lib/spreadsheet/print";
import type { PlacedPage, SheetPrint } from "./useSheetPrint";
import type { asWorkbookBook } from "./useWorkbookLoad";

/** What the extractor reads off a printed page. */
export const SHEET_EXTRACT_OPTIONS = {
  cellSelector: ".sheet-print-cell",
  clipSelector: ".sheet-print-cell",
  imageSelector: ".sheet-print-drawing img",
  canvasSelector: "canvas.sheet-print-chart, .sheet-print-drawing canvas:not(.sheet-print-chart)",
} as const;

const SETTLE_DEADLINE_MS = 8000;
export const PDF_FAILED_MESSAGE = "Couldn't export the PDF";

export interface SheetPdfExportDeps {
  book: ReturnType<typeof asWorkbookBook>;
  print: Pick<SheetPrint, "chosen" | "jobFor" | "isOpen">;
}

/** The off-screen sheet's element for a page of the job. */
function pageElement(index: number): HTMLElement | null {
  return document.querySelector<HTMLElement>(`.sheet-print-sheet [data-page-index="${index}"]`);
}

/** Waits for the page's images and the fonts, at most SETTLE_DEADLINE_MS. */
async function settle(element: HTMLElement): Promise<void> {
  const images = Array.from(element.querySelectorAll("img")).map((image) => image.decode().catch(() => undefined));
  const deadline = new Promise<void>((resolve) => setTimeout(resolve, SETTLE_DEADLINE_MS));
  await Promise.race([Promise.all([...images, document.fonts?.ready]), deadline]);
}

export function useSheetPdfExport(deps: SheetPdfExportDeps) {
  const busy = ref(false);
  const progress = ref({ done: 0, total: 0 });
  const notice = ref("");
  /** The pages mounted off-screen now (the rolling window). */
  const sheetPages = shallowRef<PlacedPage[]>([]);
  let abortRequested = false;

  function abort(): void {
    abortRequested = true;
  }

  /** Exports `pages` (default: the preview's chosen pages, else the active
   * sheet's job) as a PDF download. */
  async function exportPdf(pages?: PlacedPage[]): Promise<void> {
    const model = deps.book.workbook.value;
    if (busy.value || !model) {
      return;
    }
    const plan = pages ?? (deps.print.isOpen.value ? deps.print.chosen.value : deps.print.jobFor("sheet").pages.map((page, index) => ({ page, index })));
    notice.value = "";
    if (plan.length > MAX_SHEET_PDF_PAGES) {
      notice.value = PDF_LIMIT_MESSAGE;
      return;
    }
    if (plan.length === 0) {
      return;
    }
    busy.value = true;
    abortRequested = false;
    progress.value = { done: 0, total: plan.length };
    const sheetNames = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name]));
    const seenSheets = new Set<string>();
    const inputs: PdfPageInput[] = [];
    try {
      for (const [at, placed] of plan.entries()) {
        if (abortRequested) {
          return;
        }
        sheetPages.value = plan.slice(at, at + EXPORT_WINDOW_PAGES);
        await nextTick();
        const element = pageElement(placed.index);
        if (element === null) {
          throw new Error("the export sheet did not render");
        }
        await settle(element);
        if (abortRequested) {
          return;
        }
        const { page } = placed;
        const input = extractPageContent(element, page.paperPt.width, page.paperPt.height, SHEET_EXTRACT_OPTIONS);
        input.headings = seenSheets.has(page.sheetId) ? [] : [{ level: 1, title: sheetNames.get(page.sheetId) ?? "", yPt: 0 }];
        seenSheets.add(page.sheetId);
        inputs.push(input);
        progress.value = { done: at + 1, total: plan.length };
        await yieldToMainThread();
      }
      sheetPages.value = [];
      const [fonts, images] = await Promise.all([loadPdfFonts(inputs), loadPdfImages(inputs)]);
      const bytes = await buildPdf({ title: model.title, pages: inputs, fonts, images });
      if (abortRequested) {
        return;
      }
      triggerBlobDownload(bytes, pdfFileName(model.title, DEFAULT_SPREADSHEET_TITLE), PDF_MIME_TYPE);
    } catch (error) {
      notice.value = `${PDF_FAILED_MESSAGE} — ${error instanceof Error ? error.message : String(error)}`;
    } finally {
      sheetPages.value = [];
      busy.value = false;
      progress.value = { done: 0, total: 0 };
    }
  }

  return { busy, progress, notice, sheetPages, exportPdf, abort };
}

export type SheetPdfExport = ReturnType<typeof useSheetPdfExport>;
