// Building the assembly's PDF (SPEC-110) — usePdfAssembly's output half:
// the work behind "Save to library" and "Download" once the plan has passed
// every refusal. buildImagePdf composes a page at a time with a yield
// between, "Page {done} of {total}" is reported, and Cancel (or leaving the
// workspace) stops at the next page boundary with nothing downloaded, stored
// or changed. The bytes written are then parsed back (checkBuiltPdf): the
// stored document's PdfSummary is that very parse, and a build whose own
// output will not re-parse fails loudly and keeps nothing. Sources are only
// ever read; the result is always a NEW document, titled under SPEC-003's
// duplicate numbering.
import { onScopeDispose, ref, type Ref } from "vue";
import { listDocuments, storeMedia, uniqueTitle } from "../../lib/documentStore";
import { assembleMediaModel } from "../../lib/media/open";
import type { MediaModel } from "../../lib/media/model";
import {
  BROWSER_IMAGE_PDF_STEPS,
  buildImagePdf,
  checkBuiltPdf,
  ImagePdfCancelled,
  PAGE_NUMBER_FACE,
  type ImagePdfItem,
  type ImagePdfSheet,
  type ImagePdfSteps,
  type PdfAssemblyPage,
} from "../../lib/media/imagesToPdf";
import { PDF_MIME_TYPE, pdfFileName } from "../../lib/core/pdfBuild";
import { parsePdf, type PdfDocument } from "../../lib/core/pdfDocument";
import type { FaceKey } from "../../lib/core/pdfFonts";
import type { SfntFont } from "../../lib/core/sfnt";
import type { MediaPart } from "../../lib/core/media";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { loadPdfFonts } from "../../lib/pdfAssets";
import { DEFAULT_ASSEMBLY_TITLE } from "../../lib/assemblyPersist";

/** The browser-bound steps, injectable so the pipeline is unit-testable. */
export interface PdfAssemblySteps extends ImagePdfSteps {
  /** The bundled face page numbers are set in. */
  loadFonts: () => Promise<Map<FaceKey, SfntFont>>;
  /** The self-check's reader. */
  parse: (bytes: Uint8Array) => Promise<PdfDocument>;
  download: (bytes: Uint8Array, fileName: string, contentType: string) => void;
}

export const BROWSER_PDF_ASSEMBLY_STEPS: PdfAssemblySteps = {
  ...BROWSER_IMAGE_PDF_STEPS,
  // SPEC-034's own loader, asked for the one face a page number uses.
  loadFonts: () =>
    loadPdfFonts([
      {
        widthPt: 0,
        heightPt: 0,
        items: [{ text: "1", xPt: 0, yPt: 0, sizePt: 9, face: PAGE_NUMBER_FACE, color: [0, 0, 0] }],
        links: [],
        headings: [],
        images: [],
      },
    ]),
  parse: parsePdf,
  download: triggerBlobDownload,
};

/** Typed pass-through for steps arriving through a `.vue` prop. */
export function asPdfAssemblySteps(value: unknown): Partial<PdfAssemblySteps> | undefined {
  return value ? (value as Partial<PdfAssemblySteps>) : undefined;
}

/** Everything one build reads, taken at the moment it starts: edits made
 * while it runs do not reach the file. */
export interface PdfBuildSnapshot {
  pages: PdfAssemblyPage[];
  items: ImagePdfItem[];
  /** Each library id's current image. */
  parts: Map<string, MediaPart>;
  sheet: ImagePdfSheet;
  title: string;
  author: string;
  subject: string;
  pageNumbers: boolean;
}

export interface PdfBuildDeps {
  /** Why the PDF cannot be built ("" when it can) — checked before work. */
  refusal: () => string;
  /** The build's inputs; null while the sheet is unresolved. */
  snapshot: () => PdfBuildSnapshot | null;
  steps: PdfAssemblySteps;
}

export interface PdfBuildController {
  busy: Ref<boolean>;
  /** "Page {done} of {total}" while building ("" otherwise). */
  progress: Ref<string>;
  /** The last failure ("" when none). */
  failure: Ref<string>;
  save: () => Promise<MediaModel | null>;
  download: () => Promise<boolean>;
  cancel: () => void;
}

export function usePdfBuild(deps: PdfBuildDeps): PdfBuildController {
  const { steps } = deps;
  const busy = ref(false);
  const progress = ref("");
  const failure = ref("");
  let cancelled = false;
  let title = DEFAULT_ASSEMBLY_TITLE;

  /** The checked PDF, or null when refused or cancelled. Throws with a
   * user-facing reason when it cannot be built. */
  async function produce(): Promise<{ bytes: Uint8Array; doc: PdfDocument } | null> {
    const refusal = deps.refusal();
    const snapshot = refusal ? null : deps.snapshot();
    if (!snapshot) {
      failure.value = refusal;
      return null;
    }
    title = snapshot.title;
    cancelled = false;
    busy.value = true;
    progress.value = `Page 0 of ${snapshot.pages.length}`;
    try {
      const fonts = snapshot.pageNumbers ? await steps.loadFonts() : new Map<FaceKey, SfntFont>();
      const bytes = await buildImagePdf(snapshot.pages, snapshot.parts, {
        title: snapshot.title,
        author: snapshot.author,
        subject: snapshot.subject,
        pageNumbers: snapshot.pageNumbers,
        items: snapshot.items,
        sheet: snapshot.sheet,
        fonts,
        steps: { rasterise: steps.rasterise, between: steps.between, buildPdf: steps.buildPdf },
        cancelled: () => cancelled,
        onProgress: (done, total) => {
          progress.value = `Page ${done} of ${total}`;
        },
      });
      if (cancelled) {
        return null;
      }
      const doc = await checkBuiltPdf(bytes, snapshot.pages.length, steps.parse);
      return cancelled ? null : { bytes, doc };
    } catch (cause) {
      if (cause instanceof ImagePdfCancelled) {
        return null;
      }
      throw cause;
    } finally {
      busy.value = false;
      progress.value = "";
    }
  }

  async function run<T>(action: string, then: (built: { bytes: Uint8Array; doc: PdfDocument }) => Promise<T> | T): Promise<T | null> {
    if (busy.value) {
      return null;
    }
    failure.value = "";
    try {
      const built = await produce();
      return built ? await then(built) : null;
    } catch (cause) {
      const reason = cause instanceof Error && cause.message ? cause.message : String(cause);
      failure.value = `Could not ${action}: ${reason}`;
      return null;
    }
  }

  // Leaving the workspace mid-build stops it: no file, no record.
  onScopeDispose(() => {
    cancelled = busy.value || cancelled;
  });

  return {
    busy,
    progress,
    failure,
    save: () =>
      run("save the PDF", async ({ bytes, doc }) => {
        const taken = listDocuments().map((entry) => entry.title);
        // Its PdfSummary is the self-check's own parse of these bytes.
        const model = await assembleMediaModel(uniqueTitle(title, taken), "pdf", bytes, doc);
        storeMedia(model);
        return model;
      }),
    download: async () =>
      (await run("download the PDF", ({ bytes }) => {
        steps.download(bytes, pdfFileName(title, DEFAULT_ASSEMBLY_TITLE), PDF_MIME_TYPE);
        return true;
      })) ?? false,
    cancel: () => {
      if (busy.value) {
        cancelled = true;
      }
    },
  };
}
