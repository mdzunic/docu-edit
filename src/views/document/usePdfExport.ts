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
import type { Ref } from "vue";
import { DEFAULT_TITLE, type DocumentModel } from "../../lib/documentModel";
import { anchorCaret } from "../../lib/links";
import { buildPdf, PDF_MIME_TYPE, pdfFileName, type PdfPageInput } from "../../lib/pdf/buildPdf";
import { extractPageContent, PT_PER_CM } from "../../lib/pdf/domExtract";
import { faceFileName, type FaceKey } from "../../lib/pdf/fontMap";
import { parseSfnt, type SfntFont } from "../../lib/pdf/sfnt";
import { isDocumentContentEmpty } from "../../lib/printRange";
import type { PaginationController } from "../../composables/usePagination";

export interface PdfExportDeps {
  doc: Ref<DocumentModel | null>;
  pagination: PaginationController;
  /** The `.pages` wrapper (Element — the view's template-ref type). */
  pagesRoot: Readonly<Ref<Element | null>>;
  /** usePrintPreview's realizer: every page rendered and measured. */
  realizeAllPages: () => Promise<boolean>;
  /** Restores virtualization after the export (a no-op while the print
   * preview holds pages realized). */
  endRealizeAll: () => void;
  addToast: (message: string) => void;
}

export const PDF_EXPORT_ERROR_PREFIX = "Couldn't create the PDF";

/** Parsed font programs cached for the session — one fetch+parse per face
 * across any number of exports. Failures are NOT cached, so a transient
 * network problem does not poison later attempts. */
const faceCache = new Map<FaceKey, Promise<SfntFont | null>>();

function loadFace(face: FaceKey): Promise<SfntFont | null> {
  let pending = faceCache.get(face);
  if (!pending) {
    pending = fetch(`/fonts/${faceFileName(face)}`)
      .then(async (response) => {
        if (!response.ok) {
          throw new Error(`font request failed (${response.status})`);
        }
        return parseSfnt(new Uint8Array(await response.arrayBuffer()));
      })
      .catch(() => {
        faceCache.delete(face);
        return null;
      });
    faceCache.set(face, pending);
  }
  return pending;
}

function raf(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => resolve());
    } else {
      setTimeout(resolve, 16);
    }
  });
}

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
    const root = deps.pagesRoot.value;
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
      await raf();
    }
    return pages;
  }

  async function loadFonts(pages: PdfPageInput[]): Promise<Map<FaceKey, SfntFont>> {
    const faces = new Set<FaceKey>();
    for (const page of pages) {
      for (const item of page.items) {
        faces.add(item.face);
      }
    }
    const fonts = new Map<FaceKey, SfntFont>();
    const loaded = await Promise.all(
      Array.from(faces).map(async (face) => [face, await loadFace(face)] as const),
    );
    for (const [face, font] of loaded) {
      if (font) {
        fonts.set(face, font);
      }
    }
    if (faces.size > 0 && fonts.size === 0) {
      throw new Error("no fonts could be loaded for embedding");
    }
    return fonts;
  }

  /** "Download → PDF (.pdf)" (AC-14). */
  async function exportPdf(): Promise<void> {
    const model = deps.doc.value;
    if (!model || pdfExporting.value) {
      return;
    }
    pdfExporting.value = true;
    pdfProgress.value = { done: 0, total: 0 };
    try {
      // A content-empty document exports a zero-page PDF (AC-27) without
      // touching the DOM or fetching fonts.
      const pages = isDocumentContentEmpty(model) ? [] : await extractPages(model);
      const fonts = await loadFonts(pages);
      const bytes = buildPdf({
        // An untitled document gets empty PDF title metadata (AC-32).
        title: model.title === DEFAULT_TITLE ? "" : model.title,
        pages,
        fonts,
      });
      const url = URL.createObjectURL(new Blob([bytes], { type: PDF_MIME_TYPE }));
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = pdfFileName(model.title, DEFAULT_TITLE);
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(url), 10_000);
    } catch (error) {
      const reason =
        error instanceof Error && error.message.length > 0
          ? error.message
          : "an unexpected error occurred";
      deps.addToast(`${PDF_EXPORT_ERROR_PREFIX} — ${reason}`);
    } finally {
      pdfExporting.value = false;
      pdfProgress.value = { done: 0, total: 0 };
      deps.endRealizeAll();
    }
  }

  return { pdfExporting, pdfProgress, exportPdf };
}
