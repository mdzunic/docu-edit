// The library's search index, from the library's side (SPEC-114): which text
// each kind of entry contributes, and the pipeline that keeps the index
// current. The app layer, where the products compose — it knows a text
// document from a deck from a PDF, which the index itself (core/
// searchEngine, in its worker) never does.
//
// Extraction is per product and honest about its ceiling:
// - a text document (.docx, .odt, .md or new — every import pipeline
//   converts to a DocumentModel) through documentText;
// - a deck (.pptx, .odp) through deckText over its imported model;
// - a PDF through core/pdfText, run in the worker on the stored bytes;
// - an image has no text layer, and a workbook is not searched — each is
//   indexed by its title and listed under "Not searched" with its reason;
// - an entry whose model is not there yet (an import that has not
//   completed) is reported "import not complete" and stores nothing, so it
//   is indexed on the first run after its import finishes.
//
// The pipeline runs one document at a time, newest first, each scheduled
// off runWhenIdle (SPEC-047), the extraction of a document or deck on this
// thread and everything else in the worker — so the library renders and is
// fully usable while it runs, and every record is written as it is built.
import {
  getDocument,
  getMedia,
  getPresentation,
  listDocuments,
  listTrashedDocuments,
  type LibraryEntry,
} from "./documentStore";
import type { DocumentKind } from "./documentKinds";
import { documentText } from "./documentText";
import { deckText } from "./presentation/deckText";
import { isLegacyPresentationModel } from "./presentation/model";
import { COVERAGE_REASONS, type ExtractedText } from "./core/extractedText";
import type { BuildRequest, IndexEntry, IndexStatus } from "./core/searchEngine";
import { ensureMediaEncoded } from "./core/media";
import { runWhenIdle } from "./core/idleScheduler";
import type { LibraryIndexClient } from "./workers/libraryIndexClient";

export { INDEX_KEY_PREFIX } from "./documentStore";
export { MAX_INDEXED_CHARS } from "./core/extractedText";
export { MAX_INDEX_BYTES } from "./core/searchEngine";
export { MAX_PDF_TEXT_PAGES } from "./core/pdfText";

/** How long an idle slot may wait under load before it runs anyway. */
const IDLE_TIMEOUT_MS = 2000;

/** The words a `kind:` filter matches for each kind of entry. */
const KIND_FACETS: Record<DocumentKind, string[]> = {
  text: ["text", "document"],
  presentation: ["presentation", "deck", "slides"],
  media: ["media"],
  spreadsheet: ["spreadsheet", "workbook", "sheet"],
};

/** The `kind:` and `format:` values an entry answers to: `kind:pdf` and
 * `kind:image` reach media by what they are; `format:jpg` reaches JPEG. */
export function entryFacets(entry: Pick<LibraryEntry, "kind" | "format">): Record<string, string[]> {
  const kind = [...KIND_FACETS[entry.kind]];
  if (entry.kind === "media") {
    kind.push(entry.format === "pdf" ? "pdf" : "image");
  }
  const format: string[] = [entry.format];
  if (entry.format === "jpeg") {
    format.push("jpg");
  }
  return { kind, format };
}

/** An entry as the index sees it. */
export function toIndexEntry(entry: LibraryEntry, trashed: boolean): IndexEntry {
  return { docId: entry.id, lastModified: entry.lastModified, title: entry.title, facets: entryFacets(entry), trashed };
}

/** What a build of `entry` needs: its extracted text, a PDF's bytes for the
 * worker, or — when there is nothing to read yet — the status to report
 * without building anything. */
export function buildRequestFor(entry: LibraryEntry): BuildRequest | IndexStatus {
  const indexEntry = toIndexEntry(entry, false);
  const none = (reason: string): BuildRequest => ({ entry: indexEntry, extracted: { text: "", coverage: "none", reason } });
  const notYet: IndexStatus = { docId: entry.id, coverage: "none", reason: COVERAGE_REASONS.importIncomplete };
  const withText = (extracted: ExtractedText): BuildRequest => ({
    entry: indexEntry,
    extracted: { text: extracted.text, coverage: extracted.coverage, ...(extracted.reason ? { reason: extracted.reason } : {}) },
  });
  switch (entry.kind) {
    case "text": {
      const model = getDocument(entry.id);
      return model ? withText(documentText(model)) : notYet;
    }
    case "presentation": {
      const model = getPresentation(entry.id);
      return model && !isLegacyPresentationModel(model) ? withText(deckText(model)) : notYet;
    }
    case "media": {
      if (entry.format !== "pdf") {
        return none(COVERAGE_REASONS.image);
      }
      const model = getMedia(entry.id);
      const source = model?.media.find((part) => part.mediaId === model.sourceMediaId);
      if (!source) {
        return notYet;
      }
      ensureMediaEncoded(source);
      return { entry: indexEntry, pdfBase64: source.data };
    }
    case "spreadsheet":
      return none(COVERAGE_REASONS.spreadsheet);
    default:
      return none(COVERAGE_REASONS.unreadable);
  }
}

function isBuildRequest(value: BuildRequest | IndexStatus): value is BuildRequest {
  return "entry" in value;
}

/** How far the index has come: `done` of the library's `total` entries are
 * indexed (or known not to be searchable). */
export interface IndexProgress {
  done: number;
  total: number;
  running: boolean;
}

export interface LibraryIndexerDeps {
  client: LibraryIndexClient;
  /** Reports progress after every document. */
  onProgress: (progress: IndexProgress) => void;
  /** Reports a document's searchability as it becomes known. */
  onStatus: (status: IndexStatus) => void;
  /** Runs one slot of work; runWhenIdle by default. */
  schedule?: (work: () => void) => void;
}

/** The pipeline. `run()` reconciles the stored records with the library
 * and builds what is missing or stale; a later `run()` (or `cancel()`)
 * abandons an earlier one at its next document. */
export function createLibraryIndexer(deps: LibraryIndexerDeps) {
  const schedule = deps.schedule ?? ((work: () => void) => runWhenIdle(work, IDLE_TIMEOUT_MS));
  const idle = (): Promise<void> => new Promise((resolve) => schedule(resolve));
  let generation = 0;

  async function run(): Promise<void> {
    const current = ++generation;
    const live = listDocuments();
    const byId = new Map(live.map((entry) => [entry.id, entry]));
    const entries = [...live.map((entry) => toIndexEntry(entry, false)), ...listTrashedDocuments().map((entry) => toIndexEntry(entry, true))];
    const { pending, statuses } = await deps.client.sync(entries);
    if (current !== generation) {
      return;
    }
    statuses.forEach(deps.onStatus);
    const total = live.length;
    let done = total - pending.length;
    deps.onProgress({ done, total, running: pending.length > 0 });
    for (const docId of pending) {
      await idle();
      if (current !== generation) {
        return;
      }
      const entry = byId.get(docId);
      const request = entry ? buildRequestFor(entry) : null;
      if (request && isBuildRequest(request)) {
        try {
          const { evicted, ...status } = await deps.client.build(request);
          if (current !== generation) {
            return;
          }
          deps.onStatus(status);
          evicted.forEach(deps.onStatus);
        } catch {
          deps.onStatus({ docId, coverage: "none", reason: COVERAGE_REASONS.unreadable });
        }
      } else if (request) {
        deps.onStatus(request);
      }
      done += 1;
      deps.onProgress({ done, total, running: done < total });
    }
  }

  return {
    run,
    /** Stops a running build at its next document. */
    cancel(): void {
      generation += 1;
    },
    /** Forgets every record and builds the index again from scratch. */
    async rebuild(): Promise<void> {
      generation += 1;
      await deps.client.clear();
      await run();
    },
  };
}

export type LibraryIndexer = ReturnType<typeof createLibraryIndexer>;
