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
import {
  buildPdf,
  PDF_MIME_TYPE,
  pdfFileName,
  type PdfImageData,
  type PdfPageInput,
} from "../../lib/core/pdfBuild";
import { extractPageContent, PT_PER_CM } from "../../lib/core/pdfExtract";
import { encodeImage } from "../../lib/core/pdfImages";
import { faceFileName, type FaceKey } from "../../lib/core/pdfFonts";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { parseSfnt, type SfntFont } from "../../lib/core/sfnt";
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

/** A face load either yields the parsed program or the reason it could not
 * be used. Carrying the reason (rather than a bare null) is what lets a
 * failed export say *why* — an unreachable /fonts/ path, an HTTP status and
 * a corrupt font file are three very different things to chase. */
type FaceResult = { font: SfntFont } | { error: string };

/** Parsed font programs cached for the session — one fetch+parse per face
 * across any number of exports. Failures are NOT cached, so a transient
 * network problem does not poison later attempts. These promises never
 * reject: a failure resolves to its reason. */
const faceCache = new Map<FaceKey, Promise<FaceResult>>();

/** The bundled faces live in public/fonts/, so they are served under the
 * app's base path — not necessarily the origin root. */
function faceUrl(face: FaceKey): string {
  return `${import.meta.env.BASE_URL}fonts/${faceFileName(face)}`;
}

function loadFace(face: FaceKey): Promise<FaceResult> {
  let pending = faceCache.get(face);
  if (!pending) {
    const url = faceUrl(face);
    pending = fetch(url)
      .then(async (response) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }
        return { font: parseSfnt(new Uint8Array(await response.arrayBuffer())) };
      })
      .catch((cause: unknown) => {
        faceCache.delete(face);
        const reason = cause instanceof Error ? cause.message : String(cause);
        return { error: `${url}: ${reason}` };
      });
    faceCache.set(face, pending);
  }
  return pending;
}

/** Decodes every image the pages place, once per distinct source. An
 * image that cannot be decoded is left out rather than failing the
 * export — the rest of the document still becomes a PDF. Each source is
 * sampled for the largest size it is printed at. */
async function loadImages(pages: PdfPageInput[]): Promise<Map<string, PdfImageData>> {
  const largest = new Map<string, { widthPt: number; heightPt: number }>();
  for (const page of pages) {
    for (const item of page.images) {
      const seen = largest.get(item.src);
      largest.set(item.src, {
        widthPt: Math.max(seen?.widthPt ?? 0, item.widthPt),
        heightPt: Math.max(seen?.heightPt ?? 0, item.heightPt),
      });
    }
  }
  const images = new Map<string, PdfImageData>();
  const encoded = await Promise.all(
    Array.from(largest.entries()).map(
      async ([src, size]) =>
        [src, await encodeImage(src, size.widthPt, size.heightPt)] as const,
    ),
  );
  for (const [src, image] of encoded) {
    if (image) {
      images.set(src, image);
    }
  }
  return images;
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

  async function loadFonts(pages: PdfPageInput[]): Promise<Map<FaceKey, SfntFont>> {
    const faces = new Set<FaceKey>();
    for (const page of pages) {
      for (const item of page.items) {
        faces.add(item.face);
      }
    }
    const fonts = new Map<FaceKey, SfntFont>();
    const failures: string[] = [];
    const loaded = await Promise.all(
      Array.from(faces).map(async (face) => [face, await loadFace(face)] as const),
    );
    for (const [face, result] of loaded) {
      if ("font" in result) {
        fonts.set(face, result.font);
      } else {
        failures.push(result.error);
      }
    }
    // Only a total loss is fatal — a PDF missing one face still embeds the
    // rest. The reason rides along so the toast names what actually broke.
    if (faces.size > 0 && fonts.size === 0) {
      throw new Error(`no fonts could be loaded for embedding (${failures[0]})`);
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
      const [fonts, images] = await Promise.all([loadFonts(pages), loadImages(pages)]);
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
