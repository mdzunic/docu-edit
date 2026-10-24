// The page organiser's output (SPEC-111): the assembly workspace's PDF merge,
// active whenever the tray holds PDF pages (or a blank page). Every tray
// entry is one output page — a library PDF's page copied by core's
// assemblePdf; an image built into a one-page PDF on the current sheet by
// SPEC-110's buildImagePdf, parsed and merged as an ordinary source (no
// second code path); a blank page the same, with nothing placed on it.
//
// One flow, three destinations: "Merge to one PDF" (the whole tray, in
// order), "Extract selected pages" (the selection only) and "Save as
// separate PDFs" (one file per source-contiguous run, downloaded as a .zip
// of stored entries) — each saved to the library as a NEW SPEC-107 media
// document or downloaded. Before any work the plan is refused when it is
// empty, names a gone source, or breaks MAX_MERGE_PAGES or
// MAX_MERGE_SOURCE_BYTES; the sources are then parsed first, and one that is
// encrypted or unreadable refuses the WHOLE plan, naming the file — nothing
// partial is ever built. Copying yields every MERGE_YIELD_PAGES pages,
// reports "Page {done} of {total}", and Cancel stops at the next page
// boundary leaving nothing behind. Every produced file is parsed back
// (checkAssembledPdf) — page count and per-page geometry against the plan —
// before it is stored or downloaded; one that fails keeps nothing.
//
// The result preview is the planned output in the browser's own viewer
// (a blob: URL for SPEC-107's <object>), rebuilt after a change at most once
// every PREVIEW_DEBOUNCE_MS, never automatically past PREVIEW_MAX_PAGES, and
// always on Refresh. The drop records of the latest build feed "What
// changed". Sources are only ever read.
import { computed, onScopeDispose, ref, type ComputedRef, type Ref } from "vue";
import { listDocuments, storeMedia, uniqueTitle } from "../../lib/documentStore";
import { assembleMediaModel } from "../../lib/media/open";
import type { MediaModel } from "../../lib/media/model";
import { BROWSER_IMAGE_PDF_STEPS, ImagePdfCancelled, type ImagePdfSteps } from "../../lib/media/imagesToPdf";
import {
  assemblePdf,
  parsePdf,
  PdfAssemblyCancelled,
  PdfAssemblyRefusal,
  type PdfAssemblyHooks,
  type PdfAssemblyPlan,
  type PdfAssemblyResult,
  type PdfDocument,
  type PdfDropRecord,
} from "../../lib/core/pdfDocument";
import { PDF_MIME_TYPE, pdfFileName } from "../../lib/core/pdfBuild";
import { zip, type ZipEntryInput } from "../../lib/core/zipWriter";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { backgroundRefusal } from "../../lib/core/sheetLayout";
import type { AssemblyPlan } from "../../lib/core/assemblyLayout";
import type { SheetSetupController } from "../../composables/useSheetSetup";
import type { AssemblyOutputSettings } from "../../lib/assemblyPersist";
import {
  DEFAULT_MERGE_TITLE,
  groupDrops,
  MERGE_DESTINATIONS,
  mergeEstimateLine,
  sourceRuns,
  zipFileName,
  type DropGroup,
  type MergeDestination,
} from "../../lib/pdfMerge";
import { createMergeSources, mergeRefusal, sourceIdOf, writeMergeOutputs, type MergeOutput } from "../../lib/pdfMergeSources";
import type { TrayEntry } from "../assembly/useAssemblyTray";
import { useMergePreview, type PdfMergePreview } from "./useMergePreview";

export { PREVIEW_DEBOUNCE_MS, PREVIEW_MAX_PAGES, type PdfMergePreview } from "./useMergePreview";
export { sourceIdOf } from "../../lib/pdfMergeSources";

export const ZIP_MIME_TYPE = "application/zip";

/** The browser-bound steps, injectable so the pipeline is unit-testable. */
export interface PdfMergeSteps {
  /** SPEC-107's reader: the sources, and every output's self-check. */
  parse: (bytes: Uint8Array) => Promise<PdfDocument>;
  /** The writer (a test hands back junk to prove the self-check). */
  assemble: (plan: PdfAssemblyPlan, hooks?: PdfAssemblyHooks) => Promise<PdfAssemblyResult>;
  /** SPEC-110's steps for an image page. */
  imagePdf: ImagePdfSteps;
  /** Awaited between pages of copying. */
  between: () => Promise<void>;
  download: (bytes: Uint8Array, fileName: string, contentType: string) => void;
  zip: (entries: ZipEntryInput[]) => Promise<Uint8Array>;
  /** A blob: URL for the preview's viewer, and its release. */
  previewUrl: (bytes: Uint8Array) => string;
  revokeUrl: (url: string) => void;
  now: () => number;
}

export const BROWSER_PDF_MERGE_STEPS: PdfMergeSteps = {
  parse: parsePdf,
  assemble: assemblePdf,
  imagePdf: BROWSER_IMAGE_PDF_STEPS,
  between: yieldToMainThread,
  download: triggerBlobDownload,
  zip,
  previewUrl: (bytes) => URL.createObjectURL(new Blob([bytes as BlobPart], { type: PDF_MIME_TYPE })),
  revokeUrl: (url) => URL.revokeObjectURL(url),
  now: () => Date.now(),
};

/** Typed pass-through for steps arriving through a `.vue` prop. */
export function asPdfMergeSteps(value: unknown): Partial<PdfMergeSteps> | undefined {
  return value ? (value as Partial<PdfMergeSteps>) : undefined;
}

/** Typed pass-through for drop groups read off a `.vue` prop. */
export function asDropGroups(value: unknown): DropGroup[] {
  return Array.isArray(value) ? (value as DropGroup[]) : [];
}

export interface PdfMergeDeps {
  entries: Ref<TrayEntry[]>;
  /** The selection, in tray order: what "Extract selected pages" takes. */
  selectedKeys: ComputedRef<string[]>;
  /** Whether the tray is a PDF plan (it holds a page). */
  active: ComputedRef<boolean>;
  sheet: SheetSetupController;
  /** SPEC-109's plan: an image page's fit, align and cell colour. */
  plan: ComputedRef<AssemblyPlan>;
  output: AssemblyOutputSettings;
  steps: PdfMergeSteps;
}

export interface PdfMergeController {
  active: ComputedRef<boolean>;
  destination: Ref<MergeDestination>;
  setDestination: (value: string) => boolean;
  /** How many pages the chosen destination writes. */
  pageCount: ComputedRef<number>;
  estimate: ComputedRef<string>;
  /** Why the output cannot be produced ("" when it can). */
  refusal: ComputedRef<string>;
  canProduce: ComputedRef<boolean>;
  /** "" unless the sheet asked for a transparency a PDF cannot hold. */
  backgroundNotice: ComputedRef<string>;
  busy: Ref<boolean>;
  /** "Page {done} of {total}" while producing ("" otherwise). */
  progress: Ref<string>;
  /** The last failure ("" when none). */
  failure: Ref<string>;
  /** What the last save stored. */
  savedText: Ref<string>;
  /** The latest build's drop records, and them grouped by kind. */
  drops: Ref<PdfDropRecord[]>;
  dropGroups: ComputedRef<DropGroup[]>;
  save: () => Promise<MediaModel | null>;
  download: () => Promise<boolean>;
  cancel: () => void;
  preview: PdfMergePreview;
}

function reasonOf(cause: unknown): string {
  return cause instanceof Error && cause.message ? cause.message : String(cause);
}

function isCancellation(cause: unknown): boolean {
  return cause instanceof PdfAssemblyCancelled || cause instanceof ImagePdfCancelled;
}

export function usePdfMerge(deps: PdfMergeDeps): PdfMergeController {
  const { sheet, output, steps } = deps;
  const destination = ref<MergeDestination>("merge");
  const busy = ref(false);
  const progress = ref("");
  const failure = ref("");
  const savedText = ref("");
  const drops = ref<PdfDropRecord[]>([]);
  let cancelled = false;

  const selectedEntries = computed(() => {
    const keys = new Set(deps.selectedKeys.value);
    return deps.entries.value.filter((entry) => keys.has(entry.key));
  });
  /** The pages the chosen destination writes, in tray order. */
  const planned = computed(() => (destination.value === "extract" ? selectedEntries.value : deps.entries.value));
  const pageCount = computed(() => planned.value.length);

  // The session's sources: parses and image pages are cached.
  const sources = createMergeSources({
    parse: steps.parse,
    imagePdf: steps.imagePdf,
    sheet: () => ({ geometry: sheet.geometry.value, ...sheet.setup.value, plan: deps.plan.value }),
  });
  const metadata = () => ({ title: output.title.trim() || DEFAULT_MERGE_TITLE, author: output.author, subject: output.subject });

  const refusal = computed(() =>
    mergeRefusal(planned.value, {
      destination: destination.value,
      trayLength: deps.entries.value.length,
      sheetReady: sheet.geometry.value !== null,
    }),
  );

  const estimate = computed(() =>
    mergeEstimateLine(destination.value, {
      pages: deps.entries.value.length,
      selected: selectedEntries.value.length,
      sources: new Set(deps.entries.value.map(sourceIdOf)).size,
      runs: sourceRuns(deps.entries.value, sourceIdOf).length,
    }),
  );
  const backgroundNotice = computed(() =>
    deps.entries.value.some((entry) => !entry.page) ? backgroundRefusal(sheet.setup.value.background, "pdf") ?? "" : "",
  );
  const canProduce = computed(() => !busy.value && refusal.value === "");

  // --- producing ----------------------------------------------------------------

  /** The checked outputs of the chosen destination, or null when refused or
   * cancelled. Throws a user-facing Error when a file cannot be built. */
  async function produce(): Promise<MergeOutput[] | null> {
    const refused = refusal.value;
    if (refused) {
      failure.value = refused;
      return null;
    }
    // A snapshot: edits made while producing do not reach these files.
    const split = destination.value === "split";
    const entries = [...planned.value];
    const stop = () => cancelled;
    cancelled = false;
    busy.value = true;
    progress.value = `Page 0 of ${entries.length}`;
    try {
      const plan = await sources.prepare(entries, metadata(), stop);
      if (!plan) {
        return null;
      }
      const outputs = await writeMergeOutputs(
        { ...plan, modified: new Date() },
        {
          split,
          fileName: pdfFileName(plan.title, DEFAULT_MERGE_TITLE),
          assemble: steps.assemble,
          parse: steps.parse,
          between: steps.between,
          cancelled: stop,
          onProgress: (done) => {
            progress.value = `Page ${done} of ${entries.length}`;
          },
        },
      );
      if (outputs) {
        drops.value = outputs.flatMap((built) => built.dropped);
      }
      return outputs;
    } catch (cause) {
      if (isCancellation(cause)) {
        return null;
      }
      noteRefusal(cause);
      throw cause;
    } finally {
      busy.value = false;
      progress.value = "";
    }
  }

  /** A refused plan's records (each encrypted or unreadable source) are
   * what changed: nothing at all was built. */
  function noteRefusal(cause: unknown): void {
    if (cause instanceof PdfAssemblyRefusal && cause.dropped.length > 0) {
      drops.value = cause.dropped;
    }
  }

  async function run<T>(action: string, then: (outputs: MergeOutput[]) => Promise<T> | T): Promise<T | null> {
    if (busy.value) {
      return null;
    }
    failure.value = "";
    savedText.value = "";
    try {
      const outputs = await produce();
      return outputs ? await then(outputs) : null;
    } catch (cause) {
      failure.value = `Could not ${action}: ${reasonOf(cause)}`;
      return null;
    }
  }

  // --- the result preview -----------------------------------------------------------

  const preview = useMergePreview({
    // What the preview shows: the planned pages, and the sheet an image or
    // blank page is built on ("" while the tray is not a PDF plan).
    signature: computed(() => {
      if (!deps.active.value || planned.value.length === 0) {
        return "";
      }
      const pages = planned.value.map((entry) => [sourceIdOf(entry), entry.page?.index ?? -1, entry.rotateBy ?? 0]);
      const built = planned.value.some((entry) => !entry.page);
      return JSON.stringify({ pages, sheet: built ? [sheet.geometry.value, sheet.setup.value, deps.plan.value] : null });
    }),
    pageCount,
    build: async (stop) => {
      const plan = await sources.prepare([...planned.value], metadata(), stop);
      return plan && !stop() ? steps.assemble(plan, { between: steps.between, cancelled: stop }) : null;
    },
    onDrops: (dropped) => {
      drops.value = dropped;
    },
    previewUrl: steps.previewUrl,
    revokeUrl: steps.revokeUrl,
    now: steps.now,
    describe: (cause) => {
      if (isCancellation(cause)) {
        return null;
      }
      noteRefusal(cause);
      return reasonOf(cause);
    },
  });

  // Leaving the workspace stops a build: no file, no record.
  onScopeDispose(() => {
    cancelled = busy.value || cancelled;
  });

  function setDestination(value: string): boolean {
    const chosen = MERGE_DESTINATIONS.find((candidate) => candidate === value);
    if (!chosen) {
      return false;
    }
    destination.value = chosen;
    failure.value = "";
    return true;
  }

  return {
    active: deps.active,
    destination,
    setDestination,
    pageCount,
    estimate,
    refusal,
    canProduce,
    backgroundNotice,
    busy,
    progress,
    failure,
    savedText,
    drops,
    dropGroups: computed(() => groupDrops(drops.value)),
    save: () =>
      run("save the PDF", async (outputs) => {
        const taken = listDocuments().map((entry) => entry.title);
        const models: MediaModel[] = [];
        for (const built of outputs) {
          const title = uniqueTitle(built.title, taken);
          taken.push(title);
          // Its PdfSummary is the self-check's own parse of these bytes.
          models.push(await assembleMediaModel(title, "pdf", built.bytes, built.doc));
        }
        // Stored only once every file is built: all of them, or none.
        models.forEach(storeMedia);
        savedText.value =
          models.length === 1 ? `Saved “${models[0].title}” to the library.` : `Saved ${models.length} PDFs to the library.`;
        return models[0] ?? null;
      }),
    download: async () =>
      (await run("download the PDF", async (outputs) => {
        if (destination.value === "split") {
          const archive = await steps.zip(outputs.map((built) => ({ name: built.fileName, data: built.bytes, store: true })));
          steps.download(archive, zipFileName(output.title), ZIP_MIME_TYPE);
        } else {
          steps.download(outputs[0].bytes, outputs[0].fileName, PDF_MIME_TYPE);
        }
        return true;
      })) ?? false,
    cancel: () => {
      if (busy.value) {
        cancelled = true;
      }
    },
    preview,
  };
}
