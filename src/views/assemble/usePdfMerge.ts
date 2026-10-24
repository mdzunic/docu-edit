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
import { computed, onScopeDispose, ref, watch, type ComputedRef, type Ref } from "vue";
import { getMedia, listDocuments, storeMedia, uniqueTitle } from "../../lib/documentStore";
import { assembleMediaModel } from "../../lib/media/open";
import type { MediaModel } from "../../lib/media/model";
import { mediaPartBytes } from "../../lib/media/download";
import { currentImagePart, savedStack } from "../../lib/media/editModel";
import {
  BROWSER_IMAGE_PDF_STEPS,
  buildImagePdf,
  ImagePdfCancelled,
  planPdfPages,
  type ImagePdfItem,
  type ImagePdfSteps,
  type PdfAssemblyPage,
} from "../../lib/media/imagesToPdf";
import {
  assemblePdf,
  checkAssembledPdf,
  mergeLimitRefusal,
  openPdfSources,
  parsePdf,
  PdfAssemblyCancelled,
  plannedPdfPages,
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
import type { MediaPart } from "../../lib/core/media";
import type { SheetSetupController } from "../../composables/useSheetSetup";
import type { AssemblyOutputSettings } from "../../lib/assemblyPersist";
import {
  DEFAULT_MERGE_TITLE,
  groupDrops,
  MERGE_DESTINATIONS,
  mergeEstimateLine,
  runFileNames,
  sourceRuns,
  uniqueSourceKeys,
  zipFileName,
  type DropGroup,
  type MergeDestination,
} from "../../lib/pdfMerge";
import type { TrayEntry } from "../assembly/useAssemblyTray";

/** A change rebuilds the preview at most this often. */
export const PREVIEW_DEBOUNCE_MS = 750;
/** Past this many pages the preview is rebuilt only on request. */
export const PREVIEW_MAX_PAGES = 200;

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

export interface PdfMergePreview {
  /** The planned output's blob: URL ("" when none is built). */
  url: Ref<string>;
  busy: Ref<boolean>;
  failure: Ref<string>;
  /** The plan changed since the preview shown was built. */
  stale: Ref<boolean>;
  /** Over PREVIEW_MAX_PAGES: rebuilt only on request. */
  manual: ComputedRef<boolean>;
  /** merge-refresh: rebuild now (or as soon as the debounce allows). */
  refresh: () => void;
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

/** Which source an entry's page comes from — one key per library PDF, per
 * image, and one for every blank page. */
export function sourceIdOf(entry: Pick<TrayEntry, "id" | "page" | "blank">): string {
  return entry.blank ? "blank" : entry.page ? `pdf:${entry.id}` : `image:${entry.id}`;
}

/** The payload an entry's page is made from: a PDF's source bytes, or an
 * image's current part (its edited render, when it has edits). */
function sourcePart(entry: TrayEntry): MediaPart | undefined {
  const model = getMedia(entry.id);
  if (!model) {
    return undefined;
  }
  return entry.page ? model.media.find((part) => part.mediaId === model.sourceMediaId) : currentImagePart(model);
}

/** One output file before it is kept. */
interface MergeOutput {
  title: string;
  fileName: string;
  bytes: Uint8Array;
  doc: PdfDocument;
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

  // --- reading the sources (cached for the session: bytes never change) ---

  const bytesByMediaId = new Map<string, Uint8Array>();
  const parsed = new WeakMap<Uint8Array, Promise<PdfDocument>>();
  const imagePages = new Map<string, Promise<Uint8Array>>();

  /** Parses bytes once, however many builds ask. */
  function parseOnce(bytes: Uint8Array): Promise<PdfDocument> {
    let doc = parsed.get(bytes);
    if (!doc) {
      doc = steps.parse(bytes);
      parsed.set(bytes, doc);
    }
    return doc;
  }

  function partBytes(part: MediaPart): Uint8Array {
    let bytes = bytesByMediaId.get(part.mediaId);
    if (!bytes) {
      bytes = mediaPartBytes(part);
      bytesByMediaId.set(part.mediaId, bytes);
    }
    return bytes;
  }

  /** What an image page is built from, as one comparable string. */
  function sheetSignature(): string {
    const { background, dpi } = sheet.setup.value;
    const { fit, align, cellBackground } = deps.plan.value;
    return JSON.stringify({ geometry: sheet.geometry.value, background, dpi, fit, align, cellBackground });
  }

  /** An image (or blank) entry built into a one-page PDF on the current
   * sheet by SPEC-110's builder — cached by what built it. */
  function imagePage(entry: TrayEntry, stop: () => boolean): Promise<Uint8Array> {
    const part = entry.blank ? undefined : sourcePart(entry);
    const key = `${sourceIdOf(entry)}|${part?.mediaId ?? ""}|${sheetSignature()}`;
    let built = imagePages.get(key);
    if (!built) {
      const geometry = sheet.geometry.value;
      if (!geometry) {
        return Promise.reject(new Error("the sheet is not ready."));
      }
      const { background, dpi } = sheet.setup.value;
      let pages: PdfAssemblyPage[] = [{ itemIndices: [], placed: [] }];
      const items: ImagePdfItem[] = [];
      const parts = new Map<string, MediaPart>();
      if (!entry.blank) {
        const model = getMedia(entry.id);
        if (!model || !part) {
          return Promise.reject(new Error(`“${entry.title}” could not be read.`));
        }
        items.push({ id: entry.id, title: entry.title, widthPx: entry.widthPx, heightPx: entry.heightPx, edits: savedStack(model) });
        parts.set(entry.id, part);
        pages = planPdfPages(
          [{ id: entry.id, widthPx: entry.widthPx, heightPx: entry.heightPx }],
          { ...deps.plan.value, layout: "grid", pageMode: "one-per-page", nUp: 1 },
          geometry,
          dpi,
        );
      }
      built = buildImagePdf(pages, parts, {
        title: entry.title,
        pageNumbers: false,
        items,
        sheet: { geometry, background, dpi },
        steps: steps.imagePdf,
        cancelled: stop,
      });
      imagePages.set(key, built);
      // A failed or cancelled build is not remembered.
      built.catch(() => imagePages.delete(key));
    }
    return built;
  }

  /** The plan for `entries`: the library PDFs parsed first — one that is
   * encrypted or unreadable refuses the whole plan (openPdfSources), before
   * anything is built — then each image and blank page built and parsed.
   * Null when `stop` ended it. */
  async function prepare(entries: readonly TrayEntry[], stop: () => boolean): Promise<PdfAssemblyPlan | null> {
    const firstOf = new Map<string, TrayEntry>();
    for (const entry of entries) {
      if (!firstOf.has(sourceIdOf(entry))) {
        firstOf.set(sourceIdOf(entry), entry);
      }
    }
    const ids = Array.from(firstOf.keys());
    const keys = new Map(uniqueSourceKeys(ids.map((id) => firstOf.get(id)?.title ?? "")).map((key, i) => [ids[i], key]));
    const docs = new Map<string, PdfDocument>();
    const pdfIds = ids.filter((id) => id.startsWith("pdf:"));
    const opened = await openPdfSources(
      pdfIds.map((id) => {
        const part = sourcePart(firstOf.get(id) as TrayEntry);
        return { source: keys.get(id) ?? "", bytes: part ? partBytes(part) : new Uint8Array(0) };
      }),
      parseOnce,
    );
    pdfIds.forEach((id, i) => docs.set(id, opened[i]));
    const built = ids.filter((id) => !id.startsWith("pdf:"));
    for (const id of built) {
      if (stop()) {
        return null;
      }
      const bytes = await imagePage(firstOf.get(id) as TrayEntry, stop);
      const [doc] = await openPdfSources([{ source: keys.get(id) ?? "", bytes }], parseOnce);
      docs.set(id, doc);
    }
    if (stop()) {
      return null;
    }
    return {
      pages: entries.map((entry) => ({
        source: keys.get(sourceIdOf(entry)) ?? "",
        doc: docs.get(sourceIdOf(entry)) as PdfDocument,
        pageIndex: entry.page?.index ?? 0,
        rotateBy: entry.rotateBy ?? 0,
      })),
      title: output.title.trim() || DEFAULT_MERGE_TITLE,
      author: output.author,
      subject: output.subject,
      outline: true,
    };
  }

  // --- refusals and the estimate ------------------------------------------------

  const refusal = computed(() => {
    const entries = deps.entries.value;
    if (entries.length === 0) {
      return "Add at least one page to the tray.";
    }
    if (destination.value === "extract" && selectedEntries.value.length === 0) {
      return "Select the pages to extract: click a page, then Shift-click or Ctrl/Cmd-click to add more.";
    }
    const gone = new Set<string>();
    const sizes = new Map<string, number>();
    for (const entry of planned.value) {
      if (entry.blank) {
        continue;
      }
      const model = getMedia(entry.id);
      if (!model || !sourcePart(entry) || (!entry.page && (entry.widthPx <= 0 || entry.heightPx <= 0))) {
        gone.add(entry.title);
        continue;
      }
      sizes.set(sourceIdOf(entry), model.byteLength);
    }
    if (gone.size > 0) {
      const names = Array.from(gone, (title) => `“${title}”`).join(", ");
      return `${names} ${gone.size === 1 ? "is" : "are"} no longer in the library or could not be read; remove ${gone.size === 1 ? "it" : "them"} from the tray.`;
    }
    if (planned.value.some((entry) => !entry.page) && !sheet.geometry.value) {
      return "The sheet is not ready.";
    }
    let bytes = 0;
    for (const size of sizes.values()) {
      bytes += size;
    }
    return mergeLimitRefusal(planned.value.length, bytes) ?? "";
  });

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
    const title = output.title.trim() || DEFAULT_MERGE_TITLE;
    const stop = () => cancelled;
    cancelled = false;
    busy.value = true;
    progress.value = `Page 0 of ${entries.length}`;
    try {
      const plan = await prepare(entries, stop);
      if (!plan) {
        return null;
      }
      const modified = new Date();
      const runs = split ? sourceRuns(plan.pages, (page) => page.source) : [plan.pages];
      const names = runFileNames(runs.map((runPages) => runPages[0].source));
      const outputs: MergeOutput[] = [];
      const dropped: PdfDropRecord[] = [];
      let done = 0;
      for (const [index, runPages] of runs.entries()) {
        const runPlan: PdfAssemblyPlan = { ...plan, pages: runPages, title: split ? runPages[0].source : title, modified };
        const result = await steps.assemble(runPlan, {
          between: steps.between,
          cancelled: stop,
          onProgress: (count) => {
            progress.value = `Page ${done + count} of ${entries.length}`;
          },
        });
        if (cancelled) {
          return null;
        }
        // The self-check: what was written must read back as what was planned.
        const doc = await checkAssembledPdf(result.bytes, plannedPdfPages(runPlan), steps.parse);
        if (cancelled) {
          return null;
        }
        outputs.push({ title: runPlan.title, fileName: split ? names[index] : pdfFileName(title, DEFAULT_MERGE_TITLE), bytes: result.bytes, doc });
        dropped.push(...result.dropped);
        done += runPages.length;
      }
      drops.value = dropped;
      return outputs;
    } catch (cause) {
      if (isCancellation(cause)) {
        return null;
      }
      throw cause;
    } finally {
      busy.value = false;
      progress.value = "";
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

  // --- the result preview ---------------------------------------------------------

  const preview: PdfMergePreview = {
    url: ref(""),
    busy: ref(false),
    failure: ref(""),
    stale: ref(false),
    manual: computed(() => planned.value.length > PREVIEW_MAX_PAGES),
    refresh: () => {
      clearTimer();
      timer = setTimeout(buildPreview, Math.max(0, untilAllowed()));
    },
  };
  let timer: ReturnType<typeof setTimeout> | null = null;
  let lastBuiltAt: number | null = null;
  let previewToken = 0;
  let planVersion = 0;

  function clearTimer(): void {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  /** How long until the debounce allows the next build. */
  function untilAllowed(): number {
    return lastBuiltAt === null ? 0 : lastBuiltAt + PREVIEW_DEBOUNCE_MS - steps.now();
  }

  function showPreview(url: string): void {
    if (preview.url.value) {
      steps.revokeUrl(preview.url.value);
    }
    preview.url.value = url;
  }

  async function buildPreview(): Promise<void> {
    timer = null;
    const mine = ++previewToken;
    const version = planVersion;
    lastBuiltAt = steps.now();
    preview.busy.value = true;
    preview.failure.value = "";
    const stop = () => mine !== previewToken;
    try {
      const plan = await prepare([...planned.value], stop);
      if (!plan || stop()) {
        return;
      }
      const result = await steps.assemble(plan, { between: steps.between, cancelled: stop });
      if (stop()) {
        return;
      }
      showPreview(steps.previewUrl(result.bytes));
      drops.value = result.dropped;
      preview.stale.value = version !== planVersion;
    } catch (cause) {
      if (!stop() && !isCancellation(cause)) {
        preview.failure.value = reasonOf(cause);
      }
    } finally {
      if (!stop()) {
        preview.busy.value = false;
      }
    }
  }

  function dropPreview(): void {
    clearTimer();
    previewToken += 1;
    showPreview("");
    preview.busy.value = false;
    preview.failure.value = "";
    preview.stale.value = false;
    drops.value = [];
  }

  /** What the preview shows, as one comparable string ("" when inactive). */
  const previewSignature = computed(() => {
    if (!deps.active.value || planned.value.length === 0) {
      return "";
    }
    const pages = planned.value.map((entry) => [sourceIdOf(entry), entry.page?.index ?? -1, entry.rotateBy ?? 0]);
    return JSON.stringify({ pages, sheet: planned.value.some((entry) => !entry.page) ? sheetSignature() : null });
  });

  watch(
    previewSignature,
    (signature) => {
      planVersion += 1;
      if (!signature) {
        dropPreview();
        return;
      }
      preview.stale.value = true;
      clearTimer();
      // Never automatically past PREVIEW_MAX_PAGES: Refresh builds it.
      if (!preview.manual.value) {
        timer = setTimeout(buildPreview, Math.max(PREVIEW_DEBOUNCE_MS, untilAllowed()));
      }
    },
    { immediate: true },
  );

  // Leaving the workspace stops a build (no file, no record) and frees the
  // preview's URL.
  onScopeDispose(() => {
    cancelled = busy.value || cancelled;
    dropPreview();
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
