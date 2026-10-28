// The library search worker (SPEC-114): holds the search index and does all
// of its heavy work off the main thread — reading and writing the index
// records in IndexedDB, parsing PDFs for their text, tokenising, and
// answering searches — so indexing never blocks the app or the library.
// The engine itself is core/searchEngine.ts; this file only routes messages
// to it (the SPEC-035/036 worker pattern, over lib/workers/workerRpc.ts).
import { SearchEngine } from "../lib/core/searchEngine";
import { idbIndexStore } from "../lib/idb";
import type { LibraryIndexRequest } from "../lib/workers/libraryIndexClient";

// See wordCount.worker.ts's comment on why `self` is cast rather than the
// tsconfig gaining a "webworker" lib entry.
const ctx = self as unknown as {
  addEventListener: (type: "message", listener: (event: MessageEvent<LibraryIndexRequest & { id: number }>) => void) => void;
  postMessage: (message: unknown) => void;
};

const engine = new SearchEngine(idbIndexStore);

async function handle(request: LibraryIndexRequest): Promise<unknown> {
  switch (request.op) {
    case "sync":
      return engine.sync(request.entries);
    case "build":
      return engine.build(request.request);
    case "touch":
      return engine.touch(request.docId);
    case "search":
      return engine.search(request.request);
    case "clear":
      return engine.clear();
    case "dropStored":
      return engine.dropStored();
    default:
      return undefined;
  }
}

ctx.addEventListener("message", (event) => {
  const { id, ...request } = event.data;
  handle(request).then(
    // See wordCount.worker.ts's comment on why targetOrigin does not apply
    // to a dedicated worker's postMessage.
    // eslint-disable-next-line unicorn/require-post-message-target-origin
    (result) => ctx.postMessage({ id, result }),
    // eslint-disable-next-line unicorn/require-post-message-target-origin
    (error: unknown) => ctx.postMessage({ id, error: error instanceof Error ? error.message : String(error) }),
  );
});
