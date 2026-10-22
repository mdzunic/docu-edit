// Client-side PDF export (SPEC-034 AC-14..26, AC-32). Orchestrates the
// whole pipeline: realize every page → extract each realized page's text,
// links and headings off the live DOM (yielding to the event loop between
// pages so the UI stays responsive, AC-23) → fetch and parse the bundled
// Liberation faces the content actually uses → build the PDF → download it
// through a transient object-URL anchor, exactly like the docx/odt path.
//
// Everything is read-only over the rendered canvas and the model: a failed
// (or successful) export never mutates the document (AC-26). Failures
// surface as one toast, "Couldn't create the PDF — <reason>" (AC-24/25).
import { ref } from "vue";
import { DEFAULT_TITLE, type DocumentModel } from "../../lib/documentModel";
import { anchorCaret } from "../../lib/links";
import { buildPdf, PDF_MIME_TYPE, pdfFileName, type PdfPageInput } from "../../lib/core/pdfBuild";
import { extractPageContent, PT_PER_CM } from "../../lib/core/pdfExtract";
import { loadPdfFonts, loadPdfImages } from "../../lib/pdfAssets";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { isDocumentContentEmpty } from "../../lib/printRange";
import type { PaginationController } from "../../composables/usePagination";
import type { EditorContext } from "./context";

export interface PdfExportDeps extends EditorContext {
  pagination: PaginationController;
  /** usePrintPreview's realizer: every page rendered and measured. */
  realizeAllPages: () => Promise<boolean>;
  /** Restores virtualization after the export (a no-op while the print
   * preview holds pages realized). */
  endRealizeAll: () => void;
}

export const PDF_EXPORT_ERROR_PREFIX = "Couldn't create the PDF";

export function usePdfExport(deps: PdfExportDeps) {
  const pdfExporting = ref(false);
  const pdfProgress = ref({ done: 0, total: 0 });

  /** Extracts every realized page, yielding between pages (AC-23) and
   * reporting progress (AC-22). */
  async function extractPages(model: DocumentModel): Promise<PdfPageInput[]> {
    const realized = await deps.realizeAllPages();
    if (!realized) {
      throw new Error("the document took too long to prepare");
    }
    const root = deps.pageRef.value;
    if (!root) {
      throw new Error("the document is not on screen");
    }
    const layout = deps.pagination.pages.value;
    const pageEls = Array.from(root.querySelectorAll(".page"));
    if (pageEls.length !== layout.length) {
      throw new Error("the page layout is still settling");
    }
    pdfProgress.value = { done: 0, total: layout.length };
    const destPageForAnchor = (anchor: string) => {
      const caret = anchorCaret(model, anchor);
      return caret ? deps.pagination.pageOfCaret(caret) : null;
    };
    const pages: PdfPageInput[] = [];
    for (let index = 0; index < pageEls.length; index += 1) {
      const geometry = layout[index];
      pages.push(
        extractPageContent(
          pageEls[index] as HTMLElement,
          geometry.pageSize.widthCm * PT_PER_CM,
          geometry.pageSize.heightCm * PT_PER_CM,
          { destPageForAnchor },
        ),
      );
      pdfProgress.value = { done: index + 1, total: pageEls.length };
      // A macrotask yield, not a frame: the pages are already realized and
      // measured by this point, so there is nothing to wait for a paint for.
      // requestAnimationFrame does not fire in a background tab, so an export
      // the user tabbed away from used to stall until the deadline killed it
      // — and even in the foreground it cost a frame per page, ~4.6s of pure
      // waiting on a 277-page document.
      await yieldToMainThread();
    }
    return pages;
  }

  /** "Download → PDF (.pdf)" (AC-14). */
  async function exportPdf(): Promise<void> {
    const model = deps.doc.value;
    if (!model || pdfExporting.value) {
      return;
    }
    pdfExporting.value = true;
    pdfProgress.value = { done: 0, total: 0 };
    // Export always reads the DOM at 100% Print layout regardless of the
    // on-screen zoom or view mode (SPEC-063 AC-52): the marker's global rules
    // (App.vue) neutralise the `.pages` scale transform and restore the page
    // chrome that Web layout hides, so extractPageContent measures natural
    // page geometry that matches the untransformed computed font sizes.
    document.body.classList.add("is-exporting-pdf");
    try {
      // A content-empty document exports a zero-page PDF (AC-27) without
      // touching the DOM or fetching fonts.
      const pages = isDocumentContentEmpty(model) ? [] : await extractPages(model);
      const [fonts, images] = await Promise.all([loadPdfFonts(pages), loadPdfImages(pages)]);
      const bytes = await buildPdf({
        // An untitled document gets empty PDF title metadata (AC-32).
        title: model.title === DEFAULT_TITLE ? "" : model.title,
        pages,
        fonts,
        images,
      });
      triggerBlobDownload(bytes, pdfFileName(model.title, DEFAULT_TITLE), PDF_MIME_TYPE);
    } catch (error) {
      const reason =
        error instanceof Error && error.message.length > 0
          ? error.message
          : "an unexpected error occurred";
      deps.addToast(`${PDF_EXPORT_ERROR_PREFIX} — ${reason}`);
    } finally {
      document.body.classList.remove("is-exporting-pdf");
      pdfExporting.value = false;
      pdfProgress.value = { done: 0, total: 0 };
      deps.endRealizeAll();
    }
  }

  return { pdfExporting, pdfProgress, exportPdf };
}
