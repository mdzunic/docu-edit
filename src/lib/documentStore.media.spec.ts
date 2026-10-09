// SPEC-107 in the store: a media document's payload lives in its SPEC-024
// MediaPart, which the SPEC-047 split writes to its own `media:` record — so
// the shell record that listing, starring and renaming touch carries no
// pixels, the payload is never rewritten on a metadata change, and sizing a
// library of images never serialises one.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { buildPngFixtureBytes } from "./media/fixtures/build";

const PAGES_PDF = new Uint8Array(readFileSync(join(__dirname, "../../e2e/fixtures/spec107-pages.pdf")));

function makeIdbMock(loaded: unknown[] = []) {
  return {
    loadAllRecords: vi.fn().mockResolvedValue(loaded),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    getRecord: vi.fn().mockResolvedValue(undefined),
  };
}

async function loadStore(loaded: unknown[] = []) {
  const idb = makeIdbMock(loaded);
  vi.doMock("./idb", () => idb);
  const store = await import("./documentStore");
  const { assembleMediaModel } = await import("./media/open");
  return { store, idb, assembleMediaModel };
}

// The write-through is deferred to idle and yields a macrotask before each
// put (see documentStore.library.spec.ts): several real ticks catch up.
async function flushWrites() {
  for (let i = 0; i < 5; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

interface WrittenRecord {
  id: string;
  media?: Array<{ data: string }>;
  model?: { media: unknown[] };
}

function written(idb: ReturnType<typeof makeIdbMock>): WrittenRecord[] {
  return idb.putRecord.mock.calls.map((call) => call[0] as WrittenRecord);
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.doUnmock("./idb");
});

describe("media documents in the store (SPEC-107)", () => {
  it("writes the payload to its own media record and the shell without it", async () => {
    const { store, idb, assembleMediaModel } = await loadStore();
    const model = await assembleMediaModel("Shot", "png", await buildPngFixtureBytes());
    store.storeMedia(model);
    await flushWrites();

    const shell = written(idb).find((record) => record.id === model.id);
    const payload = written(idb).find((record) => record.id === `media:${model.id}`);
    expect(shell?.model?.media).toEqual([]);
    expect(JSON.stringify(shell)).not.toContain(model.media[0].data);
    expect(payload?.media?.[0].data).toBe(model.media[0].data);
  });

  it("never rewrites the payload on a star or a rename, and lists and sizes without serialising it", async () => {
    const { store, idb, assembleMediaModel } = await loadStore();
    const model = await assembleMediaModel("Pages", "pdf", PAGES_PDF);
    store.storeMedia(model);
    await flushWrites();
    idb.putRecord.mockClear();

    store.setStarred(model.id, true);
    expect(store.renameDocument(model.id, "Renamed")).toBe("Renamed");
    await flushWrites();
    const ids = written(idb).map((record) => record.id);
    expect(ids).toContain(model.id);
    expect(ids).not.toContain(`media:${model.id}`);

    const stringify = vi.spyOn(JSON, "stringify");
    const [entry] = store.listDocuments();
    const usage = store.libraryStorageUsage();
    expect(stringify).not.toHaveBeenCalled();
    expect(entry).toMatchObject({ id: model.id, title: "Renamed", format: "pdf", kind: "media", starred: true });
    expect(entry).not.toHaveProperty("model");
    // The size the library shows is the source's own length.
    expect(usage.byDocument).toEqual([{ id: model.id, bytes: PAGES_PDF.length }]);
  });

  it("rejoins the payload from its media record when the library loads", async () => {
    const seed = await loadStore();
    const model = await seed.assembleMediaModel("Pages", "pdf", PAGES_PDF);
    seed.store.storeMedia(model);
    await flushWrites();
    const persisted = written(seed.idb);
    vi.doUnmock("./idb");
    vi.resetModules();

    const { store } = await loadStore(persisted);
    await store.initLibrary();
    const loaded = store.getMedia(model.id);
    expect(loaded?.media[0].data).toBe(model.media[0].data);
    expect(loaded?.pdf?.pageCount).toBe(3);
    expect(store.getDocument(model.id)).toBeUndefined();
    expect(store.getPresentation(model.id)).toBeUndefined();
  });
});
