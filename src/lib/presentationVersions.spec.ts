// SPEC-103 — the deck's version history over the shared `versions` store:
// baseline/delta encode and decode, the slide window windowOf trims, blob
// refcounting with opaque∩media collapsing onto one blob, the over-budget
// refusal, pruning with re-baselining, and the list's slide deltas. The idb
// seam is mocked with an in-memory map (jsdom has no IndexedDB), as in
// versionHistory.spec.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel, Slide } from "./presentation/model";

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

/** Loads the module over a fresh fake store; `budget` lowers core's
 * HISTORY_BUDGET_BYTES so the budget paths run without 50 MB of data. */
async function load(options: { budget?: number; overrides?: Record<string, unknown>; store?: ReturnType<typeof fakeStore> } = {}) {
  const store = options.store ?? fakeStore();
  vi.doMock("./idb", () => ({
    getVersionsRecord: store.getVersionsRecord,
    putVersionsRecord: store.putVersionsRecord,
    deleteVersionsRecord: store.deleteVersionsRecord,
    ...options.overrides,
  }));
  if (options.budget !== undefined) {
    const budget = options.budget;
    vi.doMock("./core/versionChain", async (importOriginal) => ({
      ...(await importOriginal<typeof import("./core/versionChain")>()),
      HISTORY_BUDGET_BYTES: budget,
    }));
  }
  const lib = await import("./presentationVersions");
  return { lib, store };
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.doUnmock("./idb");
  vi.doUnmock("./core/versionChain");
});

function slide(n: number, extra: Partial<Slide> = {}): Slide {
  return {
    id: `slide-${n}`,
    frames: [
      {
        id: `title-${n}`,
        kind: "text",
        xCm: 1,
        yCm: 1,
        widthCm: 20,
        heightCm: 3,
        paragraphs: [{ runs: [{ text: `Slide ${n} title with a little body text to give it weight` }] }],
      } as Slide["frames"][number],
    ],
    partName: `ppt/slides/slide${n}.xml`,
    ...extra,
  };
}

/** A pptx-shaped deck: a package whose image part `media` duplicates. */
function deck(slides = 12): PresentationModel {
  return {
    kind: "presentation",
    id: "deck",
    title: "Quarterly",
    format: "pptx",
    slides: Array.from({ length: slides }, (_, i) => slide(i + 1)),
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    properties: { title: "Quarterly" },
    opaque: [
      { name: "[Content_Types].xml", data: "PFR5cGVzLz4=" },
      { name: "ppt/presentation.xml", data: "PHA6cHJlc2VudGF0aW9uLz4=" },
      { name: "ppt/media/image1.png", data: "iVBORw0KGgo=" },
    ],
    media: [{ mediaId: "ppt/media/image1.png", contentType: "image/png", data: "iVBORw0KGgo=" }],
  };
}

type Lib = Awaited<ReturnType<typeof load>>["lib"];

async function record(lib: Lib, model: PresentationModel, cachedPrev: Parameters<Lib["recordDeckVersion"]>[0]["cachedPrev"], trigger: "commit" | "cadence" | "pre-restore" | "post-restore" = "commit") {
  return lib.recordDeckVersion({ docId: "deck", model, author: "Ada", trigger, cachedPrev });
}

function payloadOf(store: ReturnType<typeof fakeStore>, versionId: string) {
  return (store.records.get(`version:deck:${versionId}`) as { payload: Record<string, unknown> }).payload;
}

function blobs(store: ReturnType<typeof fakeStore>) {
  return [...store.records.values()].filter((r) => (r as { store?: string }).store === "blob") as { hash: string; refCount: number; data: string }[];
}

describe("baseline and delta encode/decode", () => {
  it("the first version is a baseline carrying slides verbatim, and reconstructs the deck", async () => {
    const { lib, store } = await load();
    const model = deck(4);
    model.slides[1] = { ...model.slides[1], dirty: true, sourceXml: "<draw:page/>", preserved: { "p:timing": "<p:timing/>" } };
    model.structureDirty = true;
    const result = await record(lib, model, null);
    expect(result.written).toBe(true);
    const entry = result.index.versions[0];
    expect(entry).toMatchObject({ kind: "baseline", parentId: null, slides: 4, words: 0, author: "Ada", trigger: "commit" });
    const payload = payloadOf(store, entry.id);
    expect(payload.kind).toBe("deck-baseline");
    expect(payload.slides).toEqual(model.slides);
    // No bytes inline: the package and media are refs.
    expect(JSON.stringify(payload)).not.toContain("PHA6cHJlc2VudGF0aW9uLz4=");

    const back = await lib.reconstructDeckVersion("deck", entry.id);
    expect(back).toEqual(model);
  });

  it("a delta carries only the windowOf-trimmed slide window and the changed fields", async () => {
    const { lib, store } = await load();
    const v1 = deck(12);
    const first = await record(lib, v1, null);
    // An edit as commit() makes one: slide 3 re-identified, the rest kept.
    const v2: PresentationModel = { ...v1, slides: v1.slides.slice(), title: "Quarterly (final)" };
    v2.slides[2] = { ...v2.slides[2], dirty: true };
    const second = await record(lib, v2, first.newEntry);
    const entry = second.index.versions[1];
    expect(entry).toMatchObject({ kind: "delta", parentId: first.newEntry!.id });
    const payload = payloadOf(store, entry.id) as {
      slides: { prefixLen: number; suffixLen: number; replaced: Slide[] };
      fields: Record<string, unknown>;
      opaque?: unknown;
      media?: unknown;
    };
    expect(payload.slides.prefixLen).toBe(2);
    expect(payload.slides.suffixLen).toBe(9);
    expect(payload.slides.replaced).toEqual([v2.slides[2]]);
    expect(payload.fields).toEqual({ title: "Quarterly (final)" });
    // The package and media did not change, so the delta carries no refs.
    expect(payload.opaque).toBeUndefined();
    expect(payload.media).toBeUndefined();

    expect(await lib.reconstructDeckVersion("deck", entry.id)).toEqual(v2);
    expect(await lib.reconstructDeckVersion("deck", first.newEntry!.id)).toEqual(v1);
  });

  it("captures a slide a deck-wide command mutated in place (the window is keyed by content)", async () => {
    const { lib } = await load();
    const model = deck(6);
    const first = await record(lib, structuredClone(model), null);
    // Same slide objects, one of them edited in place — no re-identification.
    const live = structuredClone(model);
    const cached = await record(lib, live, first.newEntry);
    live.slides[4].hidden = true;
    const third = await record(lib, live, cached.newEntry);
    const back = await lib.reconstructDeckVersion("deck", third.newEntry!.id);
    expect(back?.slides[4].hidden).toBe(true);
  });

  it("a snapshot is the deck at the moment it was asked for — an edit landing while it writes stays out", async () => {
    const { lib } = await load();
    const live = deck(4);
    const before = structuredClone(live);
    const pending = record(lib, live, null);
    // An edit lands while the snapshot is still hashing and writing.
    (live.slides[0].frames[0] as { paragraphs: { runs: { text: string }[] }[] }).paragraphs[0].runs[0].text = "Typed later";
    live.title = "Renamed later";
    const result = await pending;
    expect(await lib.reconstructDeckVersion("deck", result.newEntry!.id)).toEqual(before);
  });

  it("a cleared field is carried as null and reconstructs absent", async () => {
    const { lib } = await load();
    const v1 = deck(12);
    const first = await record(lib, v1, null);
    const { properties: _gone, ...rest } = v1;
    const v2: PresentationModel = { ...rest };
    const second = await record(lib, v2, first.newEntry);
    const back = await lib.reconstructDeckVersion("deck", second.newEntry!.id);
    expect(back && "properties" in back).toBe(false);
    expect(back).toEqual(v2);
  });

  it("builds a delta after a reload from the stored chain, without the session's cache", async () => {
    const { lib } = await load();
    const v1 = deck(12);
    await record(lib, v1, null);
    const v2: PresentationModel = { ...v1, slides: v1.slides.slice() };
    v2.slides[0] = { ...v2.slides[0], notes: ["Speak slowly"] };
    const second = await record(lib, v2, null);
    expect(second.index.versions[1].kind).toBe("delta");
    expect(await lib.reconstructDeckVersion("deck", second.newEntry!.id)).toEqual(v2);
  });

  it("a delta carries media refs when an image was added, and reconstructs the new part", async () => {
    const { lib, store } = await load();
    const v1 = deck(12);
    const first = await record(lib, v1, null);
    const v2: PresentationModel = { ...v1, media: [...v1.media!, { mediaId: "ppt/media/image2.png", contentType: "image/png", data: "R0lGODlh" }] };
    const second = await record(lib, v2, first.newEntry);
    const payload = payloadOf(store, second.newEntry!.id) as { media?: unknown[]; opaque?: unknown };
    expect(payload.media).toHaveLength(2);
    expect(payload.opaque).toBeUndefined();
    expect((await lib.reconstructDeckVersion("deck", second.newEntry!.id))?.media).toEqual(v2.media);
  });
});

describe("blobs", () => {
  it("collapses a media part duplicating an opaque part onto one blob, counted once per payload", async () => {
    const { lib, store } = await load();
    const model = deck(12);
    const first = await record(lib, model, null);
    // Three package parts; the image's media copy shares its blob, and the
    // baseline counts it once.
    expect(blobs(store)).toHaveLength(3);
    expect(blobs(store).every((blob) => blob.refCount === 1)).toBe(true);
    // A delta that adds an image embeds the media refs again: the shared
    // image blob gains a reference, the new image gets its own blob.
    const extra = { mediaId: "ppt/media/image2.png", contentType: "image/png", data: "R0lGODlh" };
    const second = await record(lib, { ...model, media: [...model.media!, extra] }, first.newEntry);
    expect(second.index.versions[1].kind).toBe("delta");
    const counts = new Map(blobs(store).map((blob) => [blob.data, blob.refCount]));
    expect(counts.get("iVBORw0KGgo=")).toBe(2);
    expect(counts.get("R0lGODlh")).toBe(1);
    expect(counts.get("PFR5cGVzLz4=")).toBe(1);
  });

  it("restores the live deck's own package array when its content is unchanged", async () => {
    const { lib } = await load();
    const live = deck(4);
    const first = await record(lib, live, null);
    const back = await lib.reconstructDeckVersion("deck", first.newEntry!.id, live);
    expect(back?.opaque).toBe(live.opaque);
    expect(back?.media).toBe(live.media);
    // Without a live deck to lend them, the parts are read back from blobs.
    const fresh = await lib.reconstructDeckVersion("deck", first.newEntry!.id);
    expect(fresh?.opaque).not.toBe(live.opaque);
    expect(fresh?.opaque).toEqual(live.opaque);
  });

  it("is null when a blob is missing", async () => {
    const { lib, store } = await load();
    const first = await record(lib, deck(4), null);
    store.records.delete(`blob:deck:${blobs(store)[0].hash}`);
    expect(await lib.reconstructDeckVersion("deck", first.newEntry!.id)).toBeNull();
  });
});

describe("the over-budget degrade", () => {
  it("refuses a snapshot whose blob set plus baseline exceeds the budget, before any write", async () => {
    const { lib, store } = await load({ budget: 200 });
    const result = await record(lib, deck(4), null);
    expect(result.written).toBe(false);
    expect(result.index).toMatchObject({ paused: true, pausedReason: "too-large", versions: [] });
    // The index latch is the only write: no blob, no payload.
    const kinds = store.putVersionsRecord.mock.calls.map(([r]) => (r as { store: string }).store);
    expect(kinds).toEqual(["index"]);
    expect(lib.DECK_HISTORY_TOO_LARGE_MESSAGE).toBe(
      "Version history off — this deck's package is larger than the 50 MB history budget",
    );
  });

  it("stays latched while the deck is too large, even for a restore's own snapshots", async () => {
    const { lib, store } = await load({ budget: 200 });
    await record(lib, deck(4), null);
    store.putVersionsRecord.mockClear();
    const again = await record(lib, deck(4), null, "pre-restore");
    expect(again.written).toBe(false);
    expect(store.putVersionsRecord).not.toHaveBeenCalled();
  });

  it("a write failure abandons the version and never throws", async () => {
    const { lib, store } = await load({
      overrides: {
        putVersionsRecord: async (r: { id: string; store: string }) => {
          if (r.store === "payload") {
            throw new DOMException("quota", "QuotaExceededError");
          }
          store.records.set(r.id, structuredClone(r));
        },
      },
    });
    const result = await record(lib, deck(4), null);
    expect(result.written).toBe(false);
    expect(result.index.versions).toHaveLength(0);
    expect((await lib.loadDeckIndex("deck")).versions).toHaveLength(0);
  });
});

describe("pruning", () => {
  /** Twelve versions chained as deltas: each edits one more slide. */
  async function seed(lib: Lib, count: number, named = new Set<number>()) {
    let model = deck(40);
    let cachedPrev = null;
    const ids: string[] = [];
    const models: PresentationModel[] = [];
    for (let i = 0; i < count; i++) {
      if (i > 0) {
        model = { ...model, slides: model.slides.slice() };
        model.slides[i] = { ...model.slides[i], dirty: true };
      }
      const result = await record(lib, model, cachedPrev);
      cachedPrev = result.newEntry;
      ids.push(result.newEntry!.id);
      models.push(model);
      if (named.has(i)) {
        await lib.nameDeckVersion("deck", result.newEntry!.id, `Milestone ${i}`);
      }
    }
    return { ids, models };
  }

  /** Bytes the seeded history holds: payloads plus every referenced blob. */
  function used(store: ReturnType<typeof fakeStore>, index: { versions: { size: number }[] }) {
    return index.versions.reduce((sum, v) => sum + v.size, 0) + blobs(store).reduce((sum, b) => sum + b.data.length + 2, 0);
  }

  it("prunes the oldest unretained versions, re-baselining the oldest retained dependant", async () => {
    const seeded = await load();
    const { ids, models } = await seed(seeded.lib, 12);
    const budget = used(seeded.store, await seeded.lib.loadDeckIndex("deck")) - 1;
    vi.resetModules();
    // The same history, now one byte over a lowered budget.
    const { lib } = await load({ budget, store: seeded.store });
    const after = await lib.reconcileDeckBudget("deck", models[11]);
    expect(after.paused).toBe(false);
    // v0 is pruned; v1 chained only through it and goes in the same step;
    // v2 — the oldest of the ten retained — is re-encoded to stand alone.
    expect(after.versions.map((v) => v.id)).toEqual(ids.slice(2));
    expect(after.versions[0]).toMatchObject({ kind: "baseline", parentId: null });
    for (const [i, id] of ids.entries()) {
      if (i >= 2) {
        expect(await lib.reconstructDeckVersion("deck", id)).toEqual(models[i]);
      }
    }
  });

  it("latches paused (normal exhaustion) when nothing is prunable and still over", async () => {
    const seeded = await load();
    const { models } = await seed(seeded.lib, 3, new Set([0, 1, 2]));
    const budget = used(seeded.store, await seeded.lib.loadDeckIndex("deck")) - 1;
    vi.resetModules();
    const { lib } = await load({ budget, store: seeded.store });
    const index = await lib.reconcileDeckBudget("deck", models[2]);
    expect(index.paused).toBe(true);
    expect(index.pausedReason).toBeUndefined();
    expect(index.versions).toHaveLength(3);
    // A cadence snapshot is skipped while paused.
    expect((await record(lib, models[2], null, "cadence")).written).toBe(false);
  });

  it("deleteDeckVersion re-encodes the immediate child so every survivor still reconstructs", async () => {
    const { lib } = await load();
    const { ids, models } = await seed(lib, 4);
    const after = await lib.deleteDeckVersion("deck", ids[1], models[3]);
    expect(after.versions.map((v) => v.id)).toEqual([ids[0], ids[2], ids[3]]);
    expect(after.versions[1].kind).toBe("baseline");
    expect(await lib.reconstructDeckVersion("deck", ids[3])).toEqual(models[3]);
    expect(await lib.reconstructDeckVersion("deck", ids[2])).toEqual(models[2]);
  });

  it("names and unnames a version", async () => {
    const { lib } = await load();
    const { ids, models } = await seed(lib, 2);
    expect((await lib.nameDeckVersion("deck", ids[0], "  Before review ")).versions[0].name).toBe("Before review");
    const unnamed = await lib.removeDeckVersionName("deck", ids[0], models[1]);
    expect(unnamed.versions[0].name).toBeUndefined();
  });
});

describe("deckVersionEntries", () => {
  it("lists newest first with each version's slide-count delta", async () => {
    const { lib } = await load();
    const entries = lib.deckVersionEntries({
      id: "versions:deck",
      store: "index",
      docId: "deck",
      paused: false,
      versions: [
        { id: "a", timestamp: 1, author: "Ada", words: 0, slides: 4, kind: "baseline", parentId: null, size: 10 },
        { id: "b", timestamp: 2, author: "Ada", words: 0, slides: 6, kind: "delta", parentId: "a", size: 5 },
        { id: "c", timestamp: 3, author: "Ada", words: 0, slides: 5, kind: "delta", parentId: "b", size: 5 },
      ],
    });
    expect(entries.map((e) => [e.id, e.slideDelta])).toEqual([
      ["c", -1],
      ["b", 2],
      ["a", 4],
    ]);
  });
});
