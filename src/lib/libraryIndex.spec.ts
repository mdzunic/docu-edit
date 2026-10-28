// SPEC-114 — the library's search index: a cache with a stamp. Staleness
// invalidation, eviction order (by the bound and under quota pressure),
// trashed exclusion, the worker fallback when workers are unavailable, and
// the pipeline over a real library holding every kind of entry.
import { readFileSync } from "node:fs";
import { join, resolve as resolvePath } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { COVERAGE_REASONS, MAX_INDEXED_CHARS } from "./core/extractedText";
import { FakeIndexedDb } from "./core/fakeIndexedDb";
import { INDEX_VERSION } from "./core/searchIndex";
import { INDEX_FULL_REASON, MAX_INDEX_BYTES, SearchEngine, type IndexEntry, type IndexRecord, type IndexStatus, type IndexStore } from "./core/searchEngine";
import { createLibraryIndexClient } from "./workers/libraryIndexClient";

/** An in-memory IndexStore; `failPuts` refuses that many writes (quota). */
function memoryStore() {
  const map = new Map<string, IndexRecord>();
  const store = {
    map,
    failPuts: 0,
    async loadAll(): Promise<IndexRecord[]> {
      return [...map.values()].map((record) => structuredClone(record));
    },
    async put(record: IndexRecord): Promise<void> {
      if (store.failPuts > 0) {
        store.failPuts -= 1;
        throw new Error("QuotaExceededError");
      }
      map.set(record.docId, structuredClone(record));
    },
    async remove(docId: string): Promise<void> {
      map.delete(docId);
    },
    async clear(): Promise<void> {
      map.clear();
    },
  } satisfies IndexStore & { map: Map<string, IndexRecord>; failPuts: number };
  return store;
}

function entry(docId: string, lastModified: number, title = docId, trashed = false): IndexEntry {
  return { docId, lastModified, title, facets: { kind: ["text"], format: ["docx"] }, trashed };
}

function build(engine: SearchEngine, e: IndexEntry, text: string) {
  return engine.build({ entry: e, extracted: { text, coverage: "full" } });
}

const found = (engine: SearchEngine, input: string): string[] => engine.search({ input, limit: 50 }).hits.map((hit) => hit.id);

describe("staleness invalidation", () => {
  it("rebuilds a record whose lastModified moved, keeps the rest", async () => {
    const store = memoryStore();
    const first = new SearchEngine(store);
    await first.sync([entry("a", 1), entry("b", 2)]);
    await build(first, entry("a", 1), "alpha harbour");
    await build(first, entry("b", 2), "beta harbour");
    // A new session over the same storage; b was edited since.
    const second = new SearchEngine(store);
    const sync = await second.sync([entry("a", 1), entry("b", 3)]);
    expect(sync.pending).toEqual(["b"]);
    expect(sync.statuses).toEqual([{ docId: "a", coverage: "full" }]);
    // The stale record is not searchable until it is rebuilt.
    expect(found(second, "beta ")).toEqual([]);
    await build(second, entry("b", 3), "beta harbour rebuilt");
    expect(found(second, "rebuilt ")).toEqual(["b"]);
  });

  it("rebuilds every record whose INDEX_VERSION does not match", async () => {
    const store = memoryStore();
    const engine = new SearchEngine(store);
    await engine.sync([entry("a", 1)]);
    await build(engine, entry("a", 1), "alpha");
    store.map.get("a")!.indexVersion = INDEX_VERSION - 1;
    expect((await new SearchEngine(store).sync([entry("a", 1)])).pending).toEqual(["a"]);
  });

  it("deletes the record of a document that is gone", async () => {
    const store = memoryStore();
    const engine = new SearchEngine(store);
    await engine.sync([entry("a", 1), entry("gone", 1)]);
    await build(engine, entry("gone", 1), "orphan text");
    await new SearchEngine(store).sync([entry("a", 1)]);
    expect([...store.map.keys()]).toEqual([]);
  });

  it("lists pending documents newest first", async () => {
    const engine = new SearchEngine(memoryStore());
    expect((await engine.sync([entry("old", 1), entry("new", 9), entry("mid", 5)])).pending).toEqual(["new", "mid", "old"]);
  });
});

describe("trashed exclusion", () => {
  it("keeps a trashed document's record out of every search, and never builds one", async () => {
    const store = memoryStore();
    const engine = new SearchEngine(store);
    await engine.sync([entry("a", 1), entry("b", 2)]);
    await build(engine, entry("a", 1), "harbour alpha");
    await build(engine, entry("b", 2), "harbour beta");
    const sync = await engine.sync([entry("a", 1), entry("b", 2, "b", true), entry("c", 3, "c", true)]);
    expect(sync.pending).toEqual([]);
    expect(found(engine, "harbour ")).toEqual(["a"]);
    // Restored, it is searchable again without a rebuild.
    await engine.sync([entry("a", 1), entry("b", 2)]);
    expect(found(engine, "harbour ").toSorted()).toEqual(["a", "b"]);
    expect(store.map.has("c")).toBe(false);
  });
});

/** The stored size of one record over `text` (every record below holds the
 * same text, so each is this size). */
async function recordSize(text: string): Promise<number> {
  const store = memoryStore();
  const engine = new SearchEngine(store);
  await engine.sync([entry("probe", 1)]);
  await build(engine, entry("probe", 1), text);
  return store.map.get("probe")!.bytes;
}

describe("eviction order", () => {
  // Records far larger than an evicted marker, as real documents are.
  const TEXT = Array.from({ length: 400 }, (_, i) => `same size text ${i}`).join(" ");

  it("states the bound", () => {
    expect(MAX_INDEX_BYTES).toBe(64 * 1024 * 1024);
  });

  it("evicts the oldest-used records first when the library outgrows the bound", async () => {
    const size = await recordSize(TEXT);
    const store = memoryStore();
    // Room for two records, not three.
    const engine = new SearchEngine(store, { maxBytes: size * 2.5 });
    await engine.sync([entry("a", 10), entry("b", 20), entry("c", 30)]);
    await build(engine, entry("a", 10), TEXT);
    await build(engine, entry("b", 20), TEXT);
    const result = await build(engine, entry("c", 30), TEXT);
    // a was used longest ago (its lastModified): it goes.
    expect(result.evicted).toEqual([{ docId: "a", coverage: "none", reason: INDEX_FULL_REASON }]);
    expect(store.map.get("a")).toMatchObject({ evicted: true, tokens: null, reason: INDEX_FULL_REASON });
    expect(found(engine, "same ").toSorted()).toEqual(["b", "c"]);
    // An evicted document is reported, not rebuilt in a loop.
    const sync = await new SearchEngine(store, { maxBytes: size * 2.5 }).sync([entry("a", 10), entry("b", 20), entry("c", 30)]);
    expect(sync.pending).toEqual([]);
    expect(sync.statuses.find((status) => status.docId === "a")).toEqual({ docId: "a", coverage: "none", reason: INDEX_FULL_REASON });
  });

  it("counts opening a document from a search result as using it", async () => {
    const size = await recordSize(TEXT);
    const store = memoryStore();
    const engine = new SearchEngine(store, { maxBytes: size * 2.5, clock: () => 1_000 });
    await engine.sync([entry("a", 10), entry("b", 20), entry("c", 30)]);
    await build(engine, entry("a", 10), TEXT);
    await build(engine, entry("b", 20), TEXT);
    await engine.touch("a");
    expect(store.map.get("a")!.usedAt).toBe(1_000);
    const result = await build(engine, entry("c", 30), TEXT);
    expect(result.evicted.map((status) => status.docId)).toEqual(["b"]);
  });

  it("evicts the oldest-used records under quota pressure, then writes", async () => {
    const store = memoryStore();
    const engine = new SearchEngine(store);
    await engine.sync([entry("a", 10), entry("b", 20), entry("c", 30)]);
    await build(engine, entry("a", 10), "alpha words");
    await build(engine, entry("b", 20), "beta words");
    store.failPuts = 1;
    const result = await build(engine, entry("c", 30), "gamma words");
    expect(result.evicted.map((status) => status.docId)).toEqual(["a"]);
    expect(store.map.get("c")).toMatchObject({ coverage: "full" });
    expect(found(engine, "words ").toSorted()).toEqual(["b", "c"]);
  });

  it("drops every stored record when storage is under pressure, keeping this session's index", async () => {
    const store = memoryStore();
    const engine = new SearchEngine(store);
    await engine.sync([entry("a", 1)]);
    await build(engine, entry("a", 1), "alpha");
    await engine.dropStored();
    expect(store.map.size).toBe(0);
    expect(found(engine, "alpha ")).toEqual(["a"]);
  });
});

describe("bounds", () => {
  it(`cuts a document's text at MAX_INDEXED_CHARS (${MAX_INDEXED_CHARS}) and says so`, async () => {
    const engine = new SearchEngine(memoryStore());
    const long = `${"x".repeat(999)} `.repeat(Math.ceil(MAX_INDEXED_CHARS / 1000) + 5);
    await engine.sync([entry("big", 1)]);
    const result = await build(engine, entry("big", 1), `${long} tail`);
    expect(result.coverage).toBe("partial");
    expect(result.reason).toContain(COVERAGE_REASONS.sizeCeiling);
  });
});

describe("worker fallback when workers are unavailable", () => {
  it("runs the same engine in this thread", async () => {
    expect(typeof Worker).toBe("undefined");
    const store = memoryStore();
    const client = createLibraryIndexClient({ store });
    expect(client.inWorker).toBe(false);
    await client.sync([entry("a", 1, "Harbour report")]);
    await client.build({ entry: entry("a", 1, "Harbour report"), extracted: { text: "The Rotterdam figures rose.", coverage: "full" } });
    const result = await client.search({ input: '"rotterdam figures"', limit: 50 });
    expect(result.hits.map((hit) => hit.id)).toEqual(["a"]);
    expect(result.hits[0].snippets[0].filter((segment) => segment.match).map((segment) => segment.text)).toEqual(["Rotterdam figures"]);
    expect(result.hits[0].findText).toBe("rotterdam figures");
    expect(store.map.has("a")).toBe(true);
  });
});

// --- nothing leaves the browser -------------------------------------------------------

describe("nothing leaves the browser", () => {
  // SPEC-114 depends on SPEC-113 for its flag registry only: no module of
  // library search reaches the model gateway, or any network API at all
  // (SPEC-113's lib/ai/boundary.spec.ts allowlist covers the latter for
  // every module; this pins the former for this feature's own).
  const MODULES = [
    "lib/core/searchIndex.ts",
    "lib/core/searchEngine.ts",
    "lib/core/pdfText.ts",
    "lib/core/extractedText.ts",
    "lib/documentText.ts",
    "lib/presentation/deckText.ts",
    "lib/libraryIndex.ts",
    "lib/searchOpen.ts",
    "lib/findLink.ts",
    "lib/workers/libraryIndexClient.ts",
    "workers/libraryIndex.worker.ts",
    "composables/useLibrarySearch.ts",
    "views/library/LibrarySearchScope.vue",
    "views/library/LibrarySearchResults.vue",
    "views/library/LibraryFullTextSearch.vue",
    "views/preferences/LibrarySearchLabs.vue",
    "views/document/useFindLink.ts",
    "views/presentation/useDeckFindLink.ts",
  ];

  it.each(MODULES)("%s imports nothing of the model gateway and calls no network API", (module) => {
    const source = readFileSync(resolvePath(__dirname, "..", module), "utf8");
    expect(source).not.toMatch(/from "[^"]*\/ai\//);
    expect(source).not.toMatch(/useModelGateway|modelGateway|useModelConsent/);
    expect(source).not.toMatch(/\b(fetch|XMLHttpRequest|WebSocket|EventSource)\s*\(/);
  });
});

// --- the pipeline over a real library ----------------------------------------------

/** A fresh store and indexer over the installed fake IndexedDB. */
async function setup() {
  const store = await import("./documentStore");
  const libraryIndex = await import("./libraryIndex");
  const { createLibraryIndexClient: create } = await import("./workers/libraryIndexClient");
  const idb = await import("./idb");
  await store.initLibrary();
  const statuses = new Map<string, IndexStatus>();
  const progress: Array<{ done: number; total: number }> = [];
  const client = create();
  const built: string[] = [];
  const spied = { ...client, build: (request: Parameters<typeof client.build>[0]) => (built.push(request.entry.docId), client.build(request)) };
  const indexer = libraryIndex.createLibraryIndexer({
    client: spied,
    schedule: (work) => work(),
    onProgress: (next) => progress.push({ done: next.done, total: next.total }),
    onStatus: (status) => statuses.set(status.docId, status),
  });
  return { store, libraryIndex, idb, client, indexer, statuses, progress, built };
}


describe("the indexer over the library's entries", () => {
  let fake: FakeIndexedDb;
  beforeEach(() => {
    vi.resetModules();
    fake = new FakeIndexedDb();
    fake.install();
  });
  afterEach(() => {
    fake.uninstall();
  });

  it("pins the record key prefix the store deletes to the one idb excludes", async () => {
    const { libraryIndex, idb } = await setup();
    expect(libraryIndex.INDEX_KEY_PREFIX).toBe("index:");
    expect(idb.INDEX_RECORD_KEY_PREFIX).toBe(libraryIndex.INDEX_KEY_PREFIX);
  });

  it("indexes a converted .docx with coverage full, through the import pipeline's model", async () => {
    const { store, indexer, statuses, client } = await setup();
    const { readDocxBytes } = await import("./openDocument");
    const bytes = readFileSync(join(__dirname, "docx", "fixtures", "word-authored.docx"));
    const model = await readDocxBytes(bytes, "Word authored");
    store.storeDocument(model);
    await indexer.run();
    expect(statuses.get(model.id)).toEqual({ docId: model.id, coverage: "full" });
    const word = (await import("./documentText")).documentText(model).text.match(/[A-Za-z]{5,}/)?.[0] ?? "";
    expect(word).not.toBe("");
    expect((await client.search({ input: `${word} `, limit: 50 })).hits.map((hit) => hit.id)).toContain(model.id);
  });

  it("reads every kind honestly, one document at a time, newest first, reporting progress", async () => {
    const { store, indexer, statuses, progress, built, client } = await setup();
    const { createDocument } = await import("./documentModel");
    const { assembleMediaModel } = await import("./media/open");
    const { createBlankWorkbook } = await import("./spreadsheet/sheets");
    const { buildJpegFixtureBytes, buildScannedPdfBytes } = await import("./media/fixtures/build");
    const now = vi.spyOn(Date, "now");
    let clock = 1_000;
    now.mockImplementation(() => (clock += 1_000));

    const doc = createDocument();
    doc.title = "Harbour report";
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "The Rotterdam figures rose.", formatting: {} }] }];
    store.storeDocument(doc);
    const photo = await assembleMediaModel("spec107-photo", "jpeg", buildJpegFixtureBytes());
    store.storeMedia(photo);
    const scan = await assembleMediaModel("spec114-scanned", "pdf", buildScannedPdfBytes());
    store.storeMedia(scan);
    const outlines = await assembleMediaModel("spec111-outlines", "pdf", new Uint8Array(readFileSync(resolvePath(__dirname, "../../e2e/fixtures/spec111-outlines.pdf"))));
    store.storeMedia(outlines);
    const sheet = createBlankWorkbook();
    store.storeSpreadsheet(sheet);
    now.mockRestore();

    await indexer.run();
    expect(built).toEqual([sheet.id, outlines.id, scan.id, photo.id, doc.id]);
    expect(progress[0]).toEqual({ done: 0, total: 5 });
    expect(progress[progress.length - 1]).toEqual({ done: 5, total: 5 });
    expect(statuses.get(doc.id)).toEqual({ docId: doc.id, coverage: "full" });
    expect(statuses.get(outlines.id)).toEqual({ docId: outlines.id, coverage: "full" });
    expect(statuses.get(photo.id)).toEqual({ docId: photo.id, coverage: "none", reason: COVERAGE_REASONS.image });
    expect(statuses.get(scan.id)).toEqual({ docId: scan.id, coverage: "none", reason: COVERAGE_REASONS.scanned });
    expect(statuses.get(sheet.id)).toEqual({ docId: sheet.id, coverage: "none", reason: COVERAGE_REASONS.spreadsheet });

    // A PDF's text, read in the engine; an image still found by its title.
    expect((await client.search({ input: "chapter ", limit: 50 })).hits.map((hit) => hit.id)).toEqual([outlines.id]);
    // (A hyphen inside a word keeps it one token: the title is "spec107-photo".)
    expect((await client.search({ input: "spec107-photo ", limit: 50 })).hits.map((hit) => hit.id)).toEqual([photo.id]);
    // Records live under index:<id>; a second run finds them all current.
    expect([...fake.stores.get("documents")!.keys()].filter((key) => key.startsWith("index:"))).toHaveLength(5);
    built.length = 0;
    await indexer.run();
    expect(built).toEqual([]);
  });

  it("reports an import that has not completed, and stores nothing for it", async () => {
    const { store, indexer, statuses } = await setup();
    const legacy = { kind: "presentation", id: "legacy-deck", title: "Old deck", format: "pptx", slides: [{ title: "Outline", lines: [] }] };
    store.storePresentation(legacy as unknown as Parameters<typeof store.storePresentation>[0]);
    await indexer.run();
    expect(statuses.get("legacy-deck")).toEqual({ docId: "legacy-deck", coverage: "none", reason: COVERAGE_REASONS.importIncomplete });
    expect(fake.stores.get("documents")!.has("index:legacy-deck")).toBe(false);
  });

  it("deletes a document's index record with the document", async () => {
    const { store, indexer } = await setup();
    const { createDocument } = await import("./documentModel");
    const doc = createDocument();
    store.storeDocument(doc);
    await indexer.run();
    const keys = () => [...fake.stores.get("documents")!.keys()];
    await vi.waitFor(() => expect(keys()).toContain(`index:${doc.id}`));
    store.deleteDocument(doc.id);
    await vi.waitFor(() => expect(keys()).not.toContain(`index:${doc.id}`));
  });
});
