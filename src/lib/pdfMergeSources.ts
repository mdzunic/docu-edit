// The page organiser's sources and files (SPEC-111) — the framework-free
// half of the workspace's PDF merge (views/assemble/usePdfMerge wires it).
//
// Where each planned page comes from: a library PDF's stored bytes, parsed
// once however many builds ask; an image — or a blank page — built into a
// one-page PDF on the current sheet by SPEC-110's buildImagePdf and parsed
// like any other source (no second code path). The plan those pages make
// parses every library PDF FIRST, so an encrypted or unreadable one refuses
// the whole plan (openPdfSources, naming each file) before anything is
// built. The files a plan is written into — one, or one per
// source-contiguous run — are each parsed back and checked against the plan
// (checkAssembledPdf) before anything may keep them. And the refusals a plan
// meets before any work: an empty plan or selection, a source gone from the
// library, an unready sheet, and the merge bounds.
//
// Reads the store; never writes it, and never changes a source.
import { getMedia } from "./documentStore";
import { mediaPartBytes } from "./media/download";
import { currentImagePart, savedStack } from "./media/editModel";
import { buildImagePdf, planPdfPages, type ImagePdfItem, type ImagePdfSteps, type PdfAssemblyPage } from "./media/imagesToPdf";
import {
  checkAssembledPdf,
  mergeLimitRefusal,
  openPdfSources,
  plannedPdfPages,
  type PdfAssemblyHooks,
  type PdfAssemblyPlan,
  type PdfAssemblyResult,
  type PdfDocument,
  type PdfDropRecord,
} from "./core/pdfDocument";
import type { AssemblyPlan } from "./core/assemblyLayout";
import type { MediaPart } from "./core/media";
import type { SheetGeometry } from "./core/sheetLayout";
import { runFileNames, sourceRuns, uniqueSourceKeys, type MergeDestination } from "./pdfMerge";

/** One planned page, as the tray holds it. */
export interface MergeEntry {
  /** The library id of its source ("" for a blank page). */
  id: string;
  title: string;
  /** An image's pixel size (0 when unreadable). */
  widthPx: number;
  heightPx: number;
  /** A page of the library PDF `id`. */
  page?: { index: number };
  /** An empty page at the sheet's size. */
  blank?: boolean;
  /** Quarter turns added to the page. */
  rotateBy?: 0 | 90 | 180 | 270;
}

/** Which source an entry's page comes from — one key per library PDF, per
 * image, and one for every blank page. */
export function sourceIdOf(entry: Pick<MergeEntry, "id" | "page" | "blank">): string {
  return entry.blank ? "blank" : entry.page ? `pdf:${entry.id}` : `image:${entry.id}`;
}

/** The payload an entry's page is made from: a PDF's source bytes, or an
 * image's current part (its edited render, when it has edits). */
export function sourcePart(entry: Pick<MergeEntry, "id" | "page">): MediaPart | undefined {
  const model = getMedia(entry.id);
  if (!model) {
    return undefined;
  }
  return entry.page ? model.media.find((part) => part.mediaId === model.sourceMediaId) : currentImagePart(model);
}

/** The sheet an image or blank page is built on, and SPEC-109's plan (its
 * fit, align and cell colour) — read at build time. */
export interface MergeSheet {
  geometry: SheetGeometry | null;
  background: string;
  dpi: number;
  plan: AssemblyPlan;
}

export interface MergeSourcesOptions {
  parse: (bytes: Uint8Array) => Promise<PdfDocument>;
  imagePdf: ImagePdfSteps;
  sheet: () => MergeSheet;
}

/** /Info for a plan. */
export interface MergeMetadata {
  title: string;
  author?: string;
  subject?: string;
}

export interface MergeSources {
  /** The plan for `entries`, or null when `stop` ended it. Throws
   * PdfAssemblyRefusal when a source is refused. */
  prepare: (entries: readonly MergeEntry[], metadata: MergeMetadata, stop: () => boolean) => Promise<PdfAssemblyPlan | null>;
}

/** A session's sources: parses and image pages are cached (a payload's
 * bytes never change; an image page is rebuilt when its sheet does). */
export function createMergeSources(options: MergeSourcesOptions): MergeSources {
  const bytesByMediaId = new Map<string, Uint8Array>();
  const parsed = new WeakMap<Uint8Array, Promise<PdfDocument>>();
  const imagePages = new Map<string, Promise<Uint8Array>>();

  function parseOnce(bytes: Uint8Array): Promise<PdfDocument> {
    let doc = parsed.get(bytes);
    if (!doc) {
      doc = options.parse(bytes);
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

  /** An image (or blank) entry built into a one-page PDF on the current
   * sheet by SPEC-110's builder — cached by what built it. */
  function imagePage(entry: MergeEntry, stop: () => boolean): Promise<Uint8Array> {
    const sheet = options.sheet();
    const { geometry, background, dpi, plan } = sheet;
    const part = entry.blank ? undefined : sourcePart(entry);
    const signature = JSON.stringify({ geometry, background, dpi, fit: plan.fit, align: plan.align, cell: plan.cellBackground });
    const key = `${sourceIdOf(entry)}|${part?.mediaId ?? ""}|${signature}`;
    let built = imagePages.get(key);
    if (!built) {
      if (!geometry) {
        return Promise.reject(new Error("the sheet is not ready."));
      }
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
          { ...plan, layout: "grid", pageMode: "one-per-page", nUp: 1 },
          geometry,
          dpi,
        );
      }
      built = buildImagePdf(pages, parts, {
        title: entry.title,
        pageNumbers: false,
        items,
        sheet: { geometry, background, dpi },
        steps: options.imagePdf,
        cancelled: stop,
      });
      imagePages.set(key, built);
      // A failed or cancelled build is not remembered.
      built.catch(() => imagePages.delete(key));
    }
    return built;
  }

  async function prepare(entries: readonly MergeEntry[], metadata: MergeMetadata, stop: () => boolean): Promise<PdfAssemblyPlan | null> {
    const firstOf = new Map<string, MergeEntry>();
    for (const entry of entries) {
      if (!firstOf.has(sourceIdOf(entry))) {
        firstOf.set(sourceIdOf(entry), entry);
      }
    }
    const ids = Array.from(firstOf.keys());
    const keys = new Map(uniqueSourceKeys(ids.map((id) => firstOf.get(id)?.title ?? "")).map((key, i) => [ids[i], key]));
    const docs = new Map<string, PdfDocument>();
    // The library PDFs first: one that is refused refuses the plan before
    // anything is built.
    const pdfIds = ids.filter((id) => id.startsWith("pdf:"));
    const opened = await openPdfSources(
      pdfIds.map((id) => {
        const part = sourcePart(firstOf.get(id) as MergeEntry);
        return { source: keys.get(id) ?? "", bytes: part ? partBytes(part) : new Uint8Array(0) };
      }),
      parseOnce,
    );
    pdfIds.forEach((id, i) => docs.set(id, opened[i]));
    for (const id of ids.filter((candidate) => !candidate.startsWith("pdf:"))) {
      if (stop()) {
        return null;
      }
      const bytes = await imagePage(firstOf.get(id) as MergeEntry, stop);
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
      ...metadata,
      outline: true,
    };
  }

  return { prepare };
}

/** One file of the output, checked, before anything keeps it. */
export interface MergeOutput {
  title: string;
  fileName: string;
  bytes: Uint8Array;
  doc: PdfDocument;
  dropped: PdfDropRecord[];
}

export interface MergeWriteOptions {
  /** One file per source-contiguous run ("Save as separate PDFs"). */
  split: boolean;
  /** The one file's name otherwise. */
  fileName: string;
  assemble: (plan: PdfAssemblyPlan, hooks?: PdfAssemblyHooks) => Promise<PdfAssemblyResult>;
  parse: (bytes: Uint8Array) => Promise<PdfDocument>;
  between: () => Promise<void>;
  cancelled: () => boolean;
  /** Pages copied so far, across every file. */
  onProgress: (done: number) => void;
}

/** Writes `plan` into its file (or, split, one per run, each titled and
 * named for its source), parsing each back against what was planned. Null
 * when cancelled; throws when a file fails its self-check. */
export async function writeMergeOutputs(plan: PdfAssemblyPlan, options: MergeWriteOptions): Promise<MergeOutput[] | null> {
  const runs = options.split ? sourceRuns(plan.pages, (page) => page.source) : [plan.pages];
  const names = runFileNames(runs.map((pages) => pages[0].source));
  const outputs: MergeOutput[] = [];
  let done = 0;
  for (const [index, pages] of runs.entries()) {
    const runPlan: PdfAssemblyPlan = { ...plan, pages, title: options.split ? pages[0].source : plan.title };
    const result = await options.assemble(runPlan, {
      between: options.between,
      cancelled: options.cancelled,
      onProgress: (count) => options.onProgress(done + count),
    });
    if (options.cancelled()) {
      return null;
    }
    // The self-check: what was written must read back as what was planned.
    const doc = await checkAssembledPdf(result.bytes, plannedPdfPages(runPlan), options.parse);
    if (options.cancelled()) {
      return null;
    }
    outputs.push({
      title: runPlan.title,
      fileName: options.split ? names[index] : options.fileName,
      bytes: result.bytes,
      doc,
      dropped: result.dropped,
    });
    done += pages.length;
  }
  return outputs;
}

/** Why the planned pages cannot be written ("" when they can): nothing
 * planned (or, extracting, nothing selected), a source gone from the
 * library or unreadable, an unready sheet for an image or blank page, or a
 * merge bound — each before any work. */
export function mergeRefusal(
  planned: readonly MergeEntry[],
  context: { destination: MergeDestination; trayLength: number; sheetReady: boolean },
): string {
  if (context.trayLength === 0) {
    return "Add at least one page to the tray.";
  }
  if (context.destination === "extract" && planned.length === 0) {
    return "Select the pages to extract: click a page, then Shift-click or Ctrl/Cmd-click to add more.";
  }
  const gone = new Set<string>();
  const sizes = new Map<string, number>();
  for (const entry of planned) {
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
  if (planned.some((entry) => !entry.page) && !context.sheetReady) {
    return "The sheet is not ready.";
  }
  let bytes = 0;
  for (const size of sizes.values()) {
    bytes += size;
  }
  return mergeLimitRefusal(planned.length, bytes) ?? "";
}
