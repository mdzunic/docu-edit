// The library's full-text search (SPEC-114, experimental): the state behind
// the search field's scope control and the grouped results that replace the
// rows in "Everything" scope. A module-level singleton, like usePreferences —
// the scope control, the results and Labs' Rebuild/Clear share one index —
// reached only through the library-search experiment's gated surfaces, and
// inert while the flag is off: with it off nothing here creates a worker,
// builds an index or reads or writes an index record.
//
// - Titles is the default scope; switching to Everything brings the index up
//   to date (lib/libraryIndex.ts's pipeline) and searches inside documents.
// - Input is debounced SEARCH_DEBOUNCE_MS; a search still in flight when the
//   query changes is abandoned — its answer, when it comes, is dropped.
// - At most MAX_RENDERED_RESULTS documents render; "Show more" adds as many.
// - Under storage pressure (a refused write latches the store's warning) the
//   stored index records are dropped — the index is a cache.
import { computed, ref, shallowRef } from "vue";
import { isExperimentEnabled } from "../lib/experiments";
import { isStorageAvailable, onStorageStatusChange } from "../lib/documentStore";
import { createLibraryIndexer, type IndexProgress, type LibraryIndexer } from "../lib/libraryIndex";
import { createLibraryIndexClient, type LibraryIndexClient } from "../lib/workers/libraryIndexClient";
import type { IndexStatus, SearchResultHit } from "../lib/core/searchEngine";
import type { QueryError } from "../lib/core/searchIndex";

/** How long typing must pause before a search runs. */
export const SEARCH_DEBOUNCE_MS = 150;
/** How many documents the result list renders before "Show more". */
export const MAX_RENDERED_RESULTS = 50;
/** While indexing, the shown results refresh after this many documents. */
const REFRESH_EVERY = 10;

export type SearchScope = "titles" | "everything";

/** Typed empty seed for a `.vue` script's default parameter (espree bars
 * the annotation — the textEditing.noCaret idiom). */
export function noSearchHit(): SearchResultHit {
  return { id: "", score: 0, matches: 0, findText: "", title: [], snippets: [] };
}

/** The ids of a runtime-typed list of library entries. */
export function entryIds(entries: unknown): string[] {
  return (entries as Array<{ id: string }>).map((entry) => entry.id);
}

export interface LibrarySearchDeps {
  /** Whether the library-search experiment is on. */
  enabled: () => boolean;
  createClient: () => LibraryIndexClient;
  /** The indexer's idle scheduling (runWhenIdle by default). */
  schedule?: (work: () => void) => void;
}

export function createLibrarySearch(deps: LibrarySearchDeps) {
  const scope = ref<SearchScope>("titles");
  const progress = ref<IndexProgress>({ done: 0, total: 0, running: false });
  const statuses = shallowRef<Map<string, IndexStatus>>(new Map());
  const hits = shallowRef<SearchResultHit[]>([]);
  const total = ref(0);
  const error = shallowRef<QueryError | null>(null);
  const limit = ref(MAX_RENDERED_RESULTS);
  /** A search is waiting out its debounce or in flight. */
  const searching = ref(false);
  /** The query the shown results answer. */
  const answered = ref("");
  /** The last search's own duration, in ms (the SPEC-047 harness reads it). */
  const elapsedMs = ref(0);

  let client: LibraryIndexClient | null = null;
  let indexer: LibraryIndexer | null = null;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let generation = 0;
  let input = "";
  let ids: string[] | undefined;

  /** The client and pipeline, created on first need — and never while the
   * experiment is off. */
  function ready(): boolean {
    if (!deps.enabled()) {
      return false;
    }
    if (!client) {
      const created = deps.createClient();
      client = created;
      indexer = createLibraryIndexer({
        client: created,
        schedule: deps.schedule,
        onProgress: (next) => {
          const finished = progress.value.running && !next.running;
          progress.value = next;
          if (finished || (next.running && next.done % REFRESH_EVERY === 0)) {
            refresh();
          }
        },
        onStatus: (status) => {
          const next = new Map(statuses.value);
          next.set(status.docId, status);
          statuses.value = next;
        },
        // A document deleted or trashed since leaves the lists.
        onSync: (liveIds) => {
          const live = new Set(liveIds);
          statuses.value = new Map([...statuses.value].filter(([docId]) => live.has(docId)));
        },
      });
      onStorageStatusChange(() => {
        if (!isStorageAvailable()) {
          void created.dropStored().catch(() => undefined);
        }
      });
    }
    return true;
  }

  function clearResults(): void {
    hits.value = [];
    total.value = 0;
    error.value = null;
    answered.value = "";
    searching.value = false;
  }

  /** Drops the pending search and anything in flight. */
  function abandon(): void {
    generation += 1;
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }

  function run(): void {
    timer = null;
    if (!client || scope.value !== "everything" || input.trim() === "") {
      clearResults();
      return;
    }
    const current = ++generation;
    const asked = input;
    searching.value = true;
    client.search({ input: asked, limit: limit.value, ...(ids ? { ids } : {}) }).then(
      (result) => {
        if (current !== generation) {
          return;
        }
        hits.value = result.hits;
        total.value = result.total;
        error.value = result.error ?? null;
        elapsedMs.value = result.elapsedMs;
        answered.value = asked;
        searching.value = false;
      },
      () => {
        if (current === generation) {
          searching.value = false;
        }
      },
    );
  }

  /** Re-runs the current query at once (the index changed under it). */
  function refresh(): void {
    if (timer === null && answered.value !== "") {
      abandon();
      run();
    }
  }

  /** Brings the index up to date with the library — Everything scope only. */
  function refreshIndex(): void {
    if (scope.value === "everything" && ready()) {
      void indexer?.run();
    }
  }

  function setScope(next: SearchScope): void {
    scope.value = next;
    if (next === "everything") {
      refreshIndex();
      query(input, ids);
    } else {
      abandon();
      clearResults();
    }
  }

  /** Searches `text` (only `onlyIds`, when given) after the debounce. */
  function query(text: string, onlyIds?: string[]): void {
    input = text;
    ids = onlyIds;
    abandon();
    limit.value = MAX_RENDERED_RESULTS;
    if (scope.value !== "everything" || text.trim() === "" || !ready()) {
      clearResults();
      return;
    }
    searching.value = true;
    timer = setTimeout(run, SEARCH_DEBOUNCE_MS);
  }

  /** Renders MAX_RENDERED_RESULTS more documents. */
  function showMore(): void {
    limit.value += MAX_RENDERED_RESULTS;
    abandon();
    run();
  }

  /** Labs' "Rebuild index": forget every record, build from scratch. */
  async function rebuild(): Promise<void> {
    if (!ready()) {
      return;
    }
    statuses.value = new Map();
    await indexer?.rebuild();
  }

  /** Labs' "Clear index": forget every record and stop building. */
  async function clearIndex(): Promise<void> {
    if (!ready()) {
      return;
    }
    indexer?.cancel();
    await client?.clear();
    statuses.value = new Map();
    progress.value = { done: 0, total: 0, running: false };
    clearResults();
  }

  /** A result was opened: its document was just used (eviction order). */
  function opened(docId: string): void {
    void client?.touch(docId).catch(() => undefined);
  }

  /** The documents not (fully) searched, with their reasons. */
  const unsearched = computed(() => [...statuses.value.values()].filter((status) => status.coverage !== "full"));

  return {
    scope,
    progress,
    statuses,
    unsearched,
    hits,
    total,
    error,
    limit,
    searching,
    answered,
    elapsedMs,
    setScope,
    query,
    showMore,
    refreshIndex,
    rebuild,
    clearIndex,
    opened,
    abandon,
    /** Whether a client exists — false until the experiment is used. */
    started: (): boolean => client !== null,
  };
}

export type LibrarySearch = ReturnType<typeof createLibrarySearch>;

let shared: LibrarySearch | null = null;

/** The library's one search state. */
export function useLibrarySearch(): LibrarySearch {
  shared ??= createLibrarySearch({
    enabled: () => isExperimentEnabled("library-search"),
    createClient: () => createLibraryIndexClient(),
  });
  return shared;
}
