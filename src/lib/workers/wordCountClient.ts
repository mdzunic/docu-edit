// Main-thread client for the word-count worker (SPEC-047 AC-12). Falls back
// to computing inline when Worker is unavailable (Vitest/jsdom).
import type { DocumentModel } from "../documentModel";
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
 * document — exactly the block moving this off-thread is meant to avoid. */
function slimModel(model: DocumentModel): DocumentModel {
  const { media: _media, opaque: _opaque, sourceXml: _sourceXml, ...slim } = model;
  return slim as DocumentModel;
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
  const response = await getClient().call({ model: slimModel(model), includeNotes: options.includeNotes });
  return response.stats;
}
