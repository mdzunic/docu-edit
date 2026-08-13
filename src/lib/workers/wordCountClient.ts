// Main-thread client for the word-count worker (SPEC-047 AC-12). Falls back
// to computing inline when Worker is unavailable (Vitest/jsdom).
import type { Block, DocumentModel, Paragraph, Table } from "../documentModel";
import { documentStats, emptyCountStats, type CountStats } from "../wordCount";
import { createWorkerRpcClient, isWorkerSupported, type WorkerRpcClient } from "./workerRpc";

interface StatsRequest {
  model: DocumentModel;
  includeNotes: boolean;
}

interface StatsResponse {
  stats: CountStats;
}

let client: WorkerRpcClient<StatsRequest, StatsResponse> | null = null;

function getClient(): WorkerRpcClient<StatsRequest, StatsResponse> {
  if (!client) {
    client = createWorkerRpcClient<StatsRequest, StatsResponse>(
      () => new Worker(new URL("../../workers/wordCount.worker.ts", import.meta.url), { type: "module" }),
    );
  }
  return client;
}

/** The model minus its heavy media/opaque/source-xml parts, which word
 * counting never reads. Sending them anyway would make postMessage's
 * structured-clone step itself a multi-ten-MB main-thread cost on a large
 * document — exactly the block moving this off-thread is meant to avoid.
 * Used for the `includeNotes: true` path (the statistics dialog, opened
 * deliberately, not on a keystroke-driven cadence), where the projection
 * below's extra work isn't worth paying for `model.notes` too. */
function slimModel(model: DocumentModel): DocumentModel {
  const { media: _media, opaque: _opaque, sourceXml: _sourceXml, ...slim } = model;
  return slim as DocumentModel;
}

/** One block reduced to exactly what `wordCount.ts`'s `countBlock`/
 * `visibleBlockText` read (SPEC-047 AC-15's typing-latency budget): a
 * paragraph's run text and tracked-delete flag, a table's cell paragraphs,
 * and nothing at all for an image/break/toc/shape (none contribute
 * countable text — see wordCount.ts's own countBlock). Dropping
 * formatting, styleId, bookmarks, comment anchors, table borders/shading
 * and every other field cuts postMessage's structured-clone graph to a
 * small fraction of the real model's — the live count recomputes on a
 * 300ms throttle *while the user types* (AC-1's own requirement), so this
 * runs often enough that the clone's own cost matters, not just the
 * worker's compute time. wordCountClient.spec.ts pins this against the
 * real `documentStats` to catch drift if wordCount.ts ever starts reading
 * a field this projection drops. */
function projectBlockForCounting(block: Block): Block {
  if (block.type === "table") {
    const projected: Partial<Table> = {
      type: "table",
      rows: block.rows.map((row) => ({
        cells: row.cells.map((cell) => ({
          blocks: cell.blocks.map((paragraph) => projectBlockForCounting(paragraph) as Paragraph),
        })),
      })),
    };
    return projected as Table;
  }
  if (block.type === "paragraph") {
    const projected: Partial<Paragraph> = {
      type: "paragraph",
      runs: block.runs.map((run) => ({
        text: run.text,
        formatting: {},
        changes: run.changes?.delete ? { delete: run.changes.delete } : undefined,
      })),
    };
    return projected as Paragraph;
  }
  // image, break, toc, shape: countBlock contributes nothing for these —
  // only `type` is read (by isTable/isImage/isBreak/isToc's own `.type`
  // check) to route past this branch.
  return { type: block.type } as Block;
}

/** `model` with every section's blocks reduced via `projectBlockForCounting`
 * — the `includeNotes: false` path (the live count, recomputed on every
 * throttle tick while typing). `model.notes` is dropped entirely: nothing
 * downstream of an `includeNotes: false` call ever reads it. Exported only
 * for wordCountClient.spec.ts to pin against `documentStats` directly —
 * jsdom has no `Worker`, so `computeDocumentStats` itself never exercises
 * this path in a unit test (see that file's own header comment). */
export function projectModelForCounting(model: DocumentModel): DocumentModel {
  return {
    ...model,
    media: undefined,
    opaque: undefined,
    sourceXml: undefined,
    notes: [],
    sections: model.sections.map((section) => ({
      ...section,
      blocks: section.blocks.map(projectBlockForCounting),
    })),
  } as DocumentModel;
}

/** The document's word-count statistics (SPEC-036), computed in a worker so
 * a cold count over a large document never blocks the main thread. */
export async function computeDocumentStats(
  model: DocumentModel | null,
  options: { includeNotes: boolean },
): Promise<CountStats> {
  if (!model) {
    return emptyCountStats();
  }
  if (!isWorkerSupported()) {
    return documentStats(model, options);
  }
  const sent = options.includeNotes ? slimModel(model) : projectModelForCounting(model);
  const response = await getClient().call({ model: sent, includeNotes: options.includeNotes });
  return response.stats;
}
