// SPEC-114 — the library search state: scope switching, the debounce,
// abandoning an in-flight search, the 50-document render cap, and the
// flag-off path loading nothing.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SearchRequest, SearchResult } from "../lib/core/searchEngine";
import type { LibraryIndexClient } from "../lib/workers/libraryIndexClient";
import { setExperimentEnabled } from "../lib/experiments";
import { resetPreferences } from "../lib/preferences";
import { createLibrarySearch, MAX_RENDERED_RESULTS, SEARCH_DEBOUNCE_MS, useLibrarySearch } from "./useLibrarySearch";

/** A client whose searches resolve only when the test says so. */
function fakeClient() {
  const searches: Array<{ request: SearchRequest; resolve: (result: SearchResult) => void }> = [];
  const client: LibraryIndexClient = {
    inWorker: false,
    sync: vi.fn(async () => ({ pending: [], statuses: [] })),
    build: vi.fn(async () => ({ docId: "", coverage: "full" as const, evicted: [] })),
    touch: vi.fn(async () => undefined),
    search: vi.fn((request: SearchRequest) => new Promise<SearchResult>((resolve) => searches.push({ request, resolve }))),
    clear: vi.fn(async () => undefined),
    dropStored: vi.fn(async () => undefined),
    terminate: vi.fn(),
  };
  return { client, searches };
}

function result(ids: string[], total = ids.length): SearchResult {
  return { total, elapsedMs: 1, hits: ids.map((id) => ({ id, score: 1, matches: 1, findText: "x", title: [], snippets: [] })) };
}

function setup(enabled = true) {
  const { client, searches } = fakeClient();
  const createClient = vi.fn(() => client);
  const search = createLibrarySearch({ enabled: () => enabled, createClient, schedule: (work) => work() });
  return { search, client, searches, createClient };
}

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
});

describe("scope switching", () => {
  it("starts in Titles, where the field is SPEC-067's title search and nothing is indexed", () => {
    const { search, createClient } = setup();
    expect(search.scope.value).toBe("titles");
    search.query("rotterdam");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS * 2);
    expect(createClient).not.toHaveBeenCalled();
    expect(search.hits.value).toEqual([]);
  });

  it("brings the index up to date on switching to Everything, and searches inside documents", async () => {
    const { search, client, searches } = setup();
    search.query("rotterdam");
    search.setScope("everything");
    expect(client.sync).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS);
    expect(searches.map((entry) => entry.request.input)).toEqual(["rotterdam"]);
    searches[0].resolve(result(["a"]));
    await vi.runAllTimersAsync();
    expect(search.hits.value.map((hit) => hit.id)).toEqual(["a"]);
    // Back to Titles: the results go and nothing more is asked.
    search.setScope("titles");
    expect(search.hits.value).toEqual([]);
    search.query("other");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS * 2);
    expect(searches).toHaveLength(1);
  });
});

describe("debouncing", () => {
  it(`searches once typing pauses for ${SEARCH_DEBOUNCE_MS} ms`, () => {
    const { search, searches } = setup();
    search.setScope("everything");
    search.query("r");
    vi.advanceTimersByTime(100);
    search.query("ro");
    vi.advanceTimersByTime(100);
    search.query("rot");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS - 1);
    expect(searches).toHaveLength(0);
    expect(search.searching.value).toBe(true);
    vi.advanceTimersByTime(1);
    expect(searches.map((entry) => entry.request.input)).toEqual(["rot"]);
  });
});

describe("abandoning an in-flight search", () => {
  it("drops an answer that arrives after the query changed", async () => {
    const { search, searches } = setup();
    search.setScope("everything");
    search.query("rot");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS);
    search.query("rotterdam");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS);
    expect(searches.map((entry) => entry.request.input)).toEqual(["rot", "rotterdam"]);
    searches[1].resolve(result(["new"]));
    await vi.runAllTimersAsync();
    searches[0].resolve(result(["stale"]));
    await vi.runAllTimersAsync();
    expect(search.hits.value.map((hit) => hit.id)).toEqual(["new"]);
    expect(search.answered.value).toBe("rotterdam");
  });

  it("drops the answer to a query cleared while it was in flight", async () => {
    const { search, searches } = setup();
    search.setScope("everything");
    search.query("rot");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS);
    search.query("");
    searches[0].resolve(result(["stale"]));
    await vi.runAllTimersAsync();
    expect(search.hits.value).toEqual([]);
    expect(search.searching.value).toBe(false);
  });
});

describe(`the ${MAX_RENDERED_RESULTS}-document render cap`, () => {
  it("asks for at most 50 documents, and Show more for 50 more", async () => {
    const { search, searches } = setup();
    search.setScope("everything");
    search.query("report");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS);
    expect(searches[0].request.limit).toBe(50);
    const ids = Array.from({ length: 50 }, (_, i) => `d${i}`);
    searches[0].resolve(result(ids, 120));
    await vi.runAllTimersAsync();
    expect(search.hits.value).toHaveLength(50);
    expect(search.total.value).toBe(120);
    search.showMore();
    expect(searches[1].request.limit).toBe(100);
    // A new query starts over at 50.
    search.query("reports");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS);
    expect(searches[2].request.limit).toBe(50);
  });
});

describe("the flag-off path loads nothing", () => {
  it("creates no client, builds no index and searches nothing while the experiment is off", () => {
    const { search, createClient } = setup(false);
    search.setScope("everything");
    search.query("rotterdam");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS * 2);
    search.refreshIndex();
    void search.rebuild();
    expect(createClient).not.toHaveBeenCalled();
    expect(search.started()).toBe(false);
    expect(search.hits.value).toEqual([]);
  });

  it("is what the shared instance does with the library-search flag off", () => {
    localStorage.clear();
    resetPreferences();
    setExperimentEnabled("library-search", false);
    const shared = useLibrarySearch();
    shared.setScope("everything");
    shared.query("rotterdam");
    vi.advanceTimersByTime(SEARCH_DEBOUNCE_MS * 2);
    expect(shared.started()).toBe(false);
    shared.setScope("titles");
  });
});

describe("opening a result", () => {
  it("tells the index the document was used", () => {
    const { search, client } = setup();
    search.setScope("everything");
    search.opened("a");
    expect(client.touch).toHaveBeenCalledWith("a");
  });
});
