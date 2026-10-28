// SPEC-138 — the recording loop moved out of presentationVersions.ts into
// versionRecorder.ts, behind VersionCodec. Pinned here over a minimal codec
// (a model that is a list of strings) — the order of the loop's steps, the
// paused latch and the unconditional restore pair, the refusal before any
// encode, a failed write abandoning the version, blob refcounts, pruning —
// and over both product codecs sharing one store: each reconstructs only its
// own payloads. The idb seam is an in-memory map.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel } from "./presentation/model";
import type { SpreadsheetModel } from "./spreadsheet/model";

function fakeStore() {
  const records = new Map<string, unknown>();
  return {
    records,
    getVersionsRecord: vi.fn(async (id: string) => structuredClone(records.get(id))),
    putVersionsRecord: vi.fn(async (value: { id: string }) => {
      records.set(value.id, structuredClone(value));
    }),
    deleteVersionsRecord: vi.fn(async (id: string) => {
      records.delete(id);
    }),
  };
}

async function load(options: { budget?: number; store?: ReturnType<typeof fakeStore>; put?: (r: { id: string; store: string }) => Promise<void> } = {}) {
  const store = options.store ?? fakeStore();
  vi.doMock("./idb", () => ({
    getVersionsRecord: store.getVersionsRecord,
    putVersionsRecord: options.put ?? store.putVersionsRecord,
    deleteVersionsRecord: store.deleteVersionsRecord,
  }));
  if (options.budget !== undefined) {
    const budget = options.budget;
    vi.doMock("./core/versionChain", async (importOriginal) => ({
      ...(await importOriginal<typeof import("./core/versionChain")>()),
      HISTORY_BUDGET_BYTES: budget,
    }));
  }
  const recorder = await import("./versionRecorder");
  return { recorder, store };
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.doUnmock("./idb");
  vi.doUnmock("./core/versionChain");
});

/** A model of lines plus an attachment blob; deltas keep the common head. */
interface Notes {
  lines: string[];
  attachment?: string;
}
type NotesPayload =
  | { kind: "notes-baseline"; lines: string[]; attachment?: string }
  | { kind: "notes-delta"; parentId: string; keep: number; added: string[]; attachment?: string };
interface NotesTail {
  id: string;
  lines: string[];
  attachment?: string;
}

async function notesCodec(recorder: Awaited<ReturnType<typeof load>>["recorder"], log: string[] = []) {
  const { sha256Hex } = await import("./core/versionChain");
  const hashOf = (text: string) => sha256Hex(text);
  const codec: import("./versionRecorder").VersionCodec<Notes, Notes, NotesTail, NotesPayload, Notes> = {
    owns: (payload): payload is NotesPayload => payload.kind.startsWith("notes-"),
    isBaseline: (payload) => payload.kind === "notes-baseline",
    capture: (model) => {
      log.push("capture");
      return { lines: model.lines.slice(), attachment: model.attachment };
    },
    tooLargeToEncode: (capture) => capture.lines.length > 1000,
    tail: async (id, capture) => {
      log.push("tail");
      return { id, lines: capture.lines, ...(capture.attachment ? { attachment: await hashOf(capture.attachment) } : {}) };
    },
    blobs: async (capture) =>
      capture.attachment ? new Map([[await hashOf(capture.attachment), { data: capture.attachment, size: capture.attachment.length + 2 }]]) : new Map(),
    baseline: (next) => ({ kind: "notes-baseline", lines: next.lines, ...(next.attachment ? { attachment: next.attachment } : {}) }),
    delta: (prev, next) => {
      let keep = 0;
      while (keep < prev.lines.length && prev.lines[keep] === next.lines[keep]) {
        keep += 1;
      }
      return { kind: "notes-delta", parentId: prev.id, keep, added: next.lines.slice(keep), ...(next.attachment ? { attachment: next.attachment } : {}) };
    },
    stats: (next) => ({ words: next.lines.length }),
    embeddedHashes: (payload) => (payload.attachment ? [payload.attachment] : []),
    skeleton: (base) => ({ lines: (base as { lines: string[] }).lines, attachment: base.attachment }),
    apply: (parent, delta) => {
      const d = delta as Extract<NotesPayload, { kind: "notes-delta" }>;
      return { lines: [...parent.lines.slice(0, d.keep), ...d.added], attachment: d.attachment };
    },
    baselineOf: (skeleton) => ({ kind: "notes-baseline", lines: skeleton.lines, ...(skeleton.attachment ? { attachment: skeleton.attachment } : {}) }),
    tailOf: (id, skeleton) => ({ id, lines: skeleton.lines, ...(skeleton.attachment ? { attachment: skeleton.attachment } : {}) }),
    hydrate: async (_docId, skeleton, read) => ({ lines: skeleton.lines, ...(skeleton.attachment ? { attachment: await read(skeleton.attachment) } : {}) }),
    liveHashes: async (model) => new Set(model.attachment ? [await hashOf(model.attachment)] : []),
  };
  return recorder.createVersionRecorder(codec);
}

const many = (n: number, tag = "line") => Array.from({ length: n }, (_, i) => `${tag} ${i} with enough words to weigh something`);

describe("the moved loop over a minimal codec", () => {
  it("captures first, records a baseline then deltas, with the codec's stats, and reconstructs each", async () => {
    const { recorder, store } = await load();
    const log: string[] = [];
    const notes = await notesCodec(recorder, log);
    const v1: Notes = { lines: many(20), attachment: "ATTACHMENT-BYTES" };
    const first = await notes.record({ docId: "n", model: v1, author: "Ada", trigger: "commit", cachedPrev: null });
    expect(log.slice(0, 2)).toEqual(["capture", "tail"]);
    const v2: Notes = { ...v1, lines: [...v1.lines, "one more"] };
    const second = await notes.record({ docId: "n", model: v2, author: "Ada", trigger: "cadence", cachedPrev: first.newEntry });
    expect(second.index.versions.map((v) => [v.kind, v.words, v.trigger])).toEqual([
      ["baseline", 20, "commit"],
      ["delta", 21, "cadence"],
    ]);
    expect(await notes.reconstruct("n", first.newEntry!.id)).toEqual(v1);
    expect(await notes.reconstruct("n", second.newEntry!.id)).toEqual(v2);
    // The attachment is one blob, referenced by both payloads.
    const blob = [...store.records.values()].find((r) => (r as { store?: string }).store === "blob") as { refCount: number };
    expect(blob.refCount).toBe(2);
  });

  it("a snapshot is the model at the moment it was asked for", async () => {
    const { recorder } = await load();
    const notes = await notesCodec(recorder);
    const live: Notes = { lines: many(5) };
    const pending = notes.record({ docId: "n", model: live, author: "Ada", trigger: "commit", cachedPrev: null });
    live.lines.push("typed while it wrote");
    const result = await pending;
    expect((await notes.reconstruct("n", result.newEntry!.id))?.lines).toHaveLength(5);
  });

  it("refuses before any encode when the codec says so, latching too-large once", async () => {
    const { recorder, store } = await load();
    const log: string[] = [];
    const notes = await notesCodec(recorder, log);
    const result = await notes.record({ docId: "n", model: { lines: many(1001) }, author: "Ada", trigger: "commit", cachedPrev: null });
    expect(result).toMatchObject({ written: false, index: { paused: true, pausedReason: "too-large" } });
    expect(log).toEqual(["capture"]);
    store.putVersionsRecord.mockClear();
    await notes.record({ docId: "n", model: { lines: many(1001) }, author: "Ada", trigger: "post-restore", cachedPrev: null });
    expect(store.putVersionsRecord).not.toHaveBeenCalled();
  });

  it("while paused by exhaustion, cadence and commit skip but a restore's pair writes", async () => {
    const { recorder, store } = await load();
    const notes = await notesCodec(recorder);
    await store.putVersionsRecord({ id: "versions:n", store: "index", docId: "n", versions: [], paused: true } as never);
    expect((await notes.record({ docId: "n", model: { lines: many(3) }, author: "Ada", trigger: "cadence", cachedPrev: null })).written).toBe(false);
    expect((await notes.record({ docId: "n", model: { lines: many(3) }, author: "Ada", trigger: "pre-restore", cachedPrev: null })).written).toBe(true);
  });

  it("a failed payload write abandons the version without throwing", async () => {
    const store = fakeStore();
    const { recorder } = await load({
      store,
      put: async (r) => {
        if (r.store === "payload") {
          throw new DOMException("quota", "QuotaExceededError");
        }
        store.records.set(r.id, structuredClone(r));
      },
    });
    const notes = await notesCodec(recorder);
    const result = await notes.record({ docId: "n", model: { lines: many(3) }, author: "Ada", trigger: "commit", cachedPrev: null });
    expect(result.written).toBe(false);
    expect((await notes.loadIndex("n")).versions).toHaveLength(0);
  });

  it("prunes over budget with re-baselining, and latches paused when nothing is prunable", async () => {
    const seeded = await load();
    const notes = await notesCodec(seeded.recorder);
    let prev = null;
    const models: Notes[] = [];
    for (let i = 0; i < 12; i++) {
      const model = { lines: [...many(10), ...many(i, "extra")] };
      models.push(model);
      prev = (await notes.record({ docId: "n", model, author: "Ada", trigger: "commit", cachedPrev: prev })).newEntry;
    }
    const index = await notes.loadIndex("n");
    const used = index.versions.reduce((sum, v) => sum + v.size, 0);
    vi.resetModules();
    const { recorder } = await load({ budget: used - 1, store: seeded.store });
    const again = await notesCodec(recorder);
    const after = await again.reconcile("n", models[11]!);
    expect(after.paused).toBe(false);
    expect(after.versions.length).toBeLessThan(12);
    expect(after.versions[0]!.kind).toBe("baseline");
    for (const entry of after.versions) {
      expect(await again.reconstruct("n", entry.id)).not.toBeNull();
    }
  });
});

describe("both product codecs over one store", () => {
  it("a deck and a workbook record through the same loop, each reconstructing only its own payloads", async () => {
    const store = fakeStore();
    await load({ store });
    const decks = await import("./presentationVersions");
    const books = await import("./spreadsheetVersions");
    const deck: PresentationModel = {
      kind: "presentation",
      id: "deck",
      title: "Deck",
      format: "pptx",
      slideWidthCm: 25.4,
      slideHeightCm: 19.05,
      slides: [{ id: "s1", frames: [] }],
      opaque: [{ name: "[Content_Types].xml", data: "PFR5cGVzLz4=" }],
    };
    const { createBlankWorkbook } = await import("./spreadsheet/sheets");
    const book: SpreadsheetModel = { ...createBlankWorkbook(), id: "book" };
    book.sheets = [{ ...book.sheets[0]!, rows: [{ index: 0, cells: [{ col: 0, value: 1 }] }] }];
    const deckVersion = await decks.recordDeckVersion({ docId: "deck", model: deck, author: "Ada", trigger: "commit", cachedPrev: null });
    const bookVersion = await books.recordWorkbookVersion({ docId: "book", model: book, author: "Ada", trigger: "commit", cachedPrev: null });
    expect(deckVersion.index.versions[0]).toMatchObject({ slides: 1 });
    expect(bookVersion.index.versions[0]).toMatchObject({ sheets: 1, cells: 1 });
    expect(await decks.reconstructDeckVersion("deck", deckVersion.newEntry!.id)).toEqual(deck);
    expect(await books.reconstructWorkbookVersion("book", bookVersion.newEntry!.id)).toEqual(book);

    // A payload of the other kind under the same key is not this codec's.
    const deckPayload = store.records.get(`version:deck:${deckVersion.newEntry!.id}`) as { id: string };
    store.records.set(`version:book:${deckVersion.newEntry!.id}`, { ...deckPayload, id: `version:book:${deckVersion.newEntry!.id}`, docId: "book" });
    expect(await books.reconstructWorkbookVersion("book", deckVersion.newEntry!.id)).toBeNull();
  });
});
