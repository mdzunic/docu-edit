// Main-thread client for the library search worker (SPEC-114): the index
// lives in a Web Worker, reached through the shared workerRpc transport
// (the SPEC-035/036 pattern). Where workers are unavailable (Vitest/jsdom,
// a locked-down browser) the same engine runs in this thread instead —
// slower to index, identical in what it finds.
import {
  SearchEngine,
  type BuildRequest,
  type BuildResult,
  type IndexEntry,
  type IndexStore,
  type SearchRequest,
  type SearchResult,
  type SyncResult,
} from "../core/searchEngine";
import { idbIndexStore } from "../idb";
import { createWorkerRpcClient, isWorkerSupported } from "./workerRpc";

/** One message to the worker. */
export type LibraryIndexRequest =
  | { op: "sync"; entries: IndexEntry[] }
  | { op: "build"; request: BuildRequest }
  | { op: "touch"; docId: string }
  | { op: "search"; request: SearchRequest }
  | { op: "clear" }
  | { op: "dropStored" };

interface LibraryIndexResponse {
  result?: unknown;
  error?: string;
}

export interface LibraryIndexClient {
  /** Whether the index runs in a worker (false: the in-thread fallback). */
  readonly inWorker: boolean;
  sync(entries: IndexEntry[]): Promise<SyncResult>;
  build(request: BuildRequest): Promise<BuildResult>;
  touch(docId: string): Promise<void>;
  search(request: SearchRequest): Promise<SearchResult>;
  clear(): Promise<void>;
  dropStored(): Promise<void>;
  terminate(): void;
}

/** The index client: a worker where the platform has workers, else the
 * engine in this thread over `store` (IndexedDB by default). */
export function createLibraryIndexClient(options: { store?: IndexStore } = {}): LibraryIndexClient {
  if (!isWorkerSupported()) {
    const engine = new SearchEngine(options.store ?? idbIndexStore);
    return {
      inWorker: false,
      sync: (entries) => engine.sync(entries),
      build: (request) => engine.build(request),
      touch: (docId) => engine.touch(docId),
      search: async (request) => engine.search(request),
      clear: () => engine.clear(),
      dropStored: () => engine.dropStored(),
      terminate: () => undefined,
    };
  }
  const rpc = createWorkerRpcClient<LibraryIndexRequest, LibraryIndexResponse>(
    () => new Worker(new URL("../../workers/libraryIndex.worker.ts", import.meta.url), { type: "module" }),
  );
  const call = async <T>(request: LibraryIndexRequest): Promise<T> => {
    const response = await rpc.call(request);
    if (response.error !== undefined) {
      throw new Error(response.error);
    }
    return response.result as T;
  };
  return {
    inWorker: true,
    sync: (entries) => call<SyncResult>({ op: "sync", entries }),
    build: (request) => call<BuildResult>({ op: "build", request }),
    touch: (docId) => call<void>({ op: "touch", docId }),
    search: (request) => call<SearchResult>({ op: "search", request }),
    clear: () => call<void>({ op: "clear" }),
    dropStored: () => call<void>({ op: "dropStored" }),
    terminate: () => rpc.terminate(),
  };
}
