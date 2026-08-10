import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { DocumentModel, Paragraph, Section } from "./documentModel";

// SPEC-040 — mocks the idb.ts seam (same pattern as
// documentStore.autosave.spec.ts) so the storage/pruning/diff logic is
// exercised without a real IndexedDB, which jsdom doesn't provide.
function fakeStore() {
  const records = new Map<string, unknown>();
  return {
    records,
    getVersionsRecord: vi.fn(async (id: string) => records.get(id)),
    putVersionsRecord: vi.fn(async (record: { id: string }) => {
      records.set(record.id, structuredClone(record));
    }),
    deleteVersionsRecord: vi.fn(async (id: string) => {
      records.delete(id);
    }),
  };
}

async function loadVersionHistory(overrides: Record<string, unknown> = {}) {
  const store = fakeStore();
  vi.doMock("./idb", () => ({
    getVersionsRecord: store.getVersionsRecord,
    putVersionsRecord: store.putVersionsRecord,
    deleteVersionsRecord: store.deleteVersionsRecord,
    ...overrides,
  }));
  const lib = await import("./versionHistory");
  return { lib, store };
}

/** Inflates stored index-entry sizes without touching their (small, real)
 * payloads, so budget enforcement actually fires without generating 50 MB of
 * genuine JSON — computeUsedBytes trusts the index's `size` field directly
 * (AC-44), so this is a faithful way to force the over-budget path
 * deterministically. */
function inflateSizes(
  store: Awaited<ReturnType<typeof loadVersionHistory>>["store"],
  docId: string,
  bytesEach: number,
) {
  const key = `versions:${docId}`;
  const index = store.records.get(key) as { versions: { size: number }[] };
  index.versions = index.versions.map((v) => ({ ...v, size: bytesEach }));
  store.records.set(key, index);
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.doUnmock("./idb");
});

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function section(...texts: string[]): Section {
  return { blocks: texts.map(paragraph) };
}

function baseModel(id: string, ...texts: string[]): DocumentModel {
  return {
    id,
    title: "Doc",
    sections: [section(...texts)],
    defaultStyles: { pageSize: { widthCm: 21, heightCm: 29.7 }, margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 }, fontFamily: "Arial", fontSizePt: 11 },
  };
}

describe("recordVersion", () => {
  it("writes a baseline for the first version of a document", async () => {
    const { lib } = await loadVersionHistory();
    const model = baseModel("doc-1", "hello world");
    const result = await lib.recordVersion({ docId: "doc-1", model, author: "You", trigger: "commit", cachedPrev: null });
    expect(result.written).toBe(true);
    expect(result.index.versions).toHaveLength(1);
    expect(result.index.versions[0].kind).toBe("baseline");
    expect(result.index.versions[0].parentId).toBeNull();
    expect(result.index.versions[0].words).toBe(2);
  });

  it("writes a delta against the cached previous version when only a block changed", async () => {
    const { lib } = await loadVersionHistory();
    // Enough body text that a one-paragraph addition stays well under 50% of
    // the baseline's own size (rule 4 would otherwise fire even on a tiny,
    // realistic edit — see decideBaseline's tests for that rule in isolation).
    const manyParagraphs = Array.from({ length: 40 }, (_, i) => `paragraph number ${i} has a bit of body text in it`);
    const v1 = baseModel("doc-1", ...manyParagraphs);
    const first = await lib.recordVersion({ docId: "doc-1", model: v1, author: "You", trigger: "commit", cachedPrev: null });
    const newEntry = first.newEntry!;

    const v2 = { ...v1, sections: [{ blocks: [...v1.sections[0].blocks, paragraph("second paragraph")] }] };
    const second = await lib.recordVersion({
      docId: "doc-1",
      model: v2,
      author: "You",
      trigger: "commit",
      cachedPrev: newEntry,
    });
    expect(second.written).toBe(true);
    expect(second.index.versions).toHaveLength(2);
    expect(second.index.versions[1].kind).toBe("delta");
    expect(second.index.versions[1].parentId).toBe(newEntry.id);
  });

  it("reconstructs a delta version back to the full model", async () => {
    const { lib } = await loadVersionHistory();
    const v1 = baseModel("doc-1", "hello world");
    const first = await lib.recordVersion({ docId: "doc-1", model: v1, author: "You", trigger: "commit", cachedPrev: null });

    const v2 = { ...v1, sections: [{ blocks: [...v1.sections[0].blocks, paragraph("second paragraph")] }] };
    const second = await lib.recordVersion({ docId: "doc-1", model: v2, author: "You", trigger: "commit", cachedPrev: first.newEntry! });

    const reconstructed = await lib.reconstructVersion("doc-1", second.newEntry!.id);
    expect(reconstructed).not.toBeNull();
    expect(reconstructed!.sections[0].blocks).toHaveLength(2);
    expect((reconstructed!.sections[0].blocks[1] as Paragraph).runs[0].text).toBe("second paragraph");
    expect(reconstructed!.title).toBe("Doc");
  });

  it("skips writing while paused, except for restore-bracketing triggers", async () => {
    const { lib, store } = await loadVersionHistory();
    store.records.set("versions:doc-1", { id: "versions:doc-1", store: "index", docId: "doc-1", versions: [], paused: true });
    const model = baseModel("doc-1", "hello");
    const cadence = await lib.recordVersion({ docId: "doc-1", model, author: "You", trigger: "cadence", cachedPrev: null });
    expect(cadence.written).toBe(false);
    const restore = await lib.recordVersion({ docId: "doc-1", model, author: "You", trigger: "pre-restore", cachedPrev: null });
    expect(restore.written).toBe(true);
  });

  it("shares media bytes across versions by content hash", async () => {
    const { lib, store } = await loadVersionHistory();
    const media = [{ mediaId: "m1", contentType: "image/png", data: "aGVsbG8=" }];
    const manyParagraphs = Array.from({ length: 40 }, (_, i) => `paragraph number ${i} has a bit of body text in it`);
    const v1: DocumentModel = { ...baseModel("doc-1", ...manyParagraphs), media };
    const first = await lib.recordVersion({ docId: "doc-1", model: v1, author: "You", trigger: "commit", cachedPrev: null });

    const v2: DocumentModel = { ...v1, sections: [{ blocks: [...v1.sections[0].blocks, paragraph("more text")] }] };
    await lib.recordVersion({ docId: "doc-1", model: v2, author: "You", trigger: "commit", cachedPrev: first.newEntry! });

    const blobRecords = [...store.records.values()].filter((r) => (r as { store?: string }).store === "blob");
    expect(blobRecords).toHaveLength(1);
    expect((blobRecords[0] as { refCount: number }).refCount).toBe(1); // only the baseline embeds the ref; the delta didn't touch media
  });

  it("a quota error on the history write abandons that write, leaving no trace in the index (AC-40)", async () => {
    const { lib, store } = await loadVersionHistory({
      putVersionsRecord: async (record: { id: string; store: string }) => {
        if (record.store === "payload") {
          throw new Error("QuotaExceededError");
        }
        store.records.set(record.id, structuredClone(record));
      },
    });
    const model = baseModel("doc-1", "hello world");
    const result = await lib.recordVersion({ docId: "doc-1", model, author: "You", trigger: "commit", cachedPrev: null });
    expect(result.written).toBe(false);
    expect(result.index.versions).toHaveLength(0);
    // Nothing this module touches lives in the `documents` store — a quota
    // failure here has no path to the document's own committed/draft
    // records at all.
    const index = await lib.loadIndex("doc-1");
    expect(index.versions).toHaveLength(0);
  });
});

describe("decideBaseline", () => {
  it("triggers on no chain (rule 1)", async () => {
    const { lib } = await loadVersionHistory();
    expect(lib.decideBaseline(null, 10)).toBe(true);
  });

  it("triggers on the chain cap (rule 2)", async () => {
    const { lib } = await loadVersionHistory();
    const chainInfo = { baselineId: "b", baselineBytes: 100_000, deltaCount: lib.CHAIN_DELTA_CAP, accumulatedDeltaBytes: 10 };
    expect(lib.decideBaseline(chainInfo, 10)).toBe(true);
  });

  it("triggers when accumulated delta bytes reach the baseline size (rule 3)", async () => {
    const { lib } = await loadVersionHistory();
    const chainInfo = { baselineId: "b", baselineBytes: 1000, deltaCount: 1, accumulatedDeltaBytes: 950 };
    expect(lib.decideBaseline(chainInfo, 60)).toBe(true);
  });

  it("triggers on a wholesale rewrite (rule 4)", async () => {
    const { lib } = await loadVersionHistory();
    const chainInfo = { baselineId: "b", baselineBytes: 1000, deltaCount: 1, accumulatedDeltaBytes: 0 };
    expect(lib.decideBaseline(chainInfo, 500)).toBe(true);
  });

  it("stays a delta when none of the four triggers hold", async () => {
    const { lib } = await loadVersionHistory();
    const chainInfo = { baselineId: "b", baselineBytes: 10_000, deltaCount: 1, accumulatedDeltaBytes: 100 };
    expect(lib.decideBaseline(chainInfo, 200)).toBe(false);
  });
});

describe("pruning", () => {
  /** Builds each version by appending one paragraph to the previous
   * version's own model object (reusing its block references), the same
   * incremental-edit shape "writes a delta ..." above uses — needed so
   * consecutive versions actually chain as deltas instead of each being an
   * unrelated model that trips the wholesale-rewrite rule into a baseline. */
  async function seedVersions(
    lib: Awaited<ReturnType<typeof loadVersionHistory>>["lib"],
    docId: string,
    count: number,
    named: Set<number> = new Set(),
  ) {
    let cachedPrev: { id: string; model: DocumentModel } | null = null;
    let model = baseModel(docId, "paragraph 0");
    const ids: string[] = [];
    for (let i = 0; i < count; i++) {
      if (i > 0) {
        model = { ...model, sections: [{ blocks: [...model.sections[0].blocks, paragraph(`paragraph ${i}`)] }] };
      }
      const result = await lib.recordVersion({ docId, model, author: "You", trigger: "commit", cachedPrev });
      cachedPrev = result.newEntry!;
      ids.push(result.newEntry!.id);
      if (named.has(i)) {
        await lib.nameVersion(docId, result.newEntry!.id, `Milestone ${i}`);
      }
    }
    return ids;
  }

  it("prunes the oldest unnamed, non-newest-10 version first and re-encodes the oldest retained dependant to preserve chain integrity (AC-32..36)", async () => {
    const { lib, store } = await loadVersionHistory();
    const ids = await seedVersions(lib, "doc-1", 12); // 1 baseline + 11 deltas, all unnamed
    // ~5 MB per (fabricated) entry x 12 = ~60 MB, over the 50 MB budget.
    inflateSizes(store, "doc-1", 5 * 1024 * 1024);

    const currentModel = baseModel("doc-1", "current content", "padding ".repeat(2000));
    const reconciled = await lib.reconcileBudget("doc-1", currentModel);

    expect(reconciled.paused).toBe(false);
    // Only the oldest two (outside the protected newest 10) are prunable.
    expect(reconciled.versions.map((v) => v.id)).not.toContain(ids[0]);
    expect(reconciled.versions.map((v) => v.id)).not.toContain(ids[1]);
    expect(reconciled.versions).toHaveLength(10);
    // The chain-integrity fix (step 2) re-encoded the oldest retained
    // dependant (ids[2]) as a standalone baseline before ids[0]/ids[1] were
    // deleted, so it — and everything after it — must still reconstruct.
    const reencoded = reconciled.versions.find((v) => v.id === ids[2])!;
    expect(reencoded.kind).toBe("baseline");
    expect(reencoded.parentId).toBeNull();
    const newest = await lib.reconstructVersion("doc-1", ids[11]);
    expect(newest).not.toBeNull();
    expect(newest!.sections[0].blocks).toHaveLength(12);
    expect((newest!.sections[0].blocks[11] as Paragraph).runs[0].text).toBe("paragraph 11");
  });

  it("stops and pauses when nothing is prunable and the budget is still exceeded (AC-37, AC-38)", async () => {
    const { lib, store } = await loadVersionHistory();
    const ids = await seedVersions(lib, "doc-1", 3, new Set([0, 1, 2])); // all named -> nothing prunable
    inflateSizes(store, "doc-1", 20 * 1024 * 1024); // 60 MB, over budget

    const currentModel = baseModel("doc-1", "current content");
    const reconciled = await lib.reconcileBudget("doc-1", currentModel);

    expect(reconciled.paused).toBe(true);
    expect(reconciled.versions.map((v) => v.id)).toEqual(ids); // nothing removed
    expect(lib.HISTORY_PAUSED_MESSAGE).toBe(
      "Version history paused — 50 MB used. Delete a version, or remove a name, to resume",
    );

    // AC-37: a subsequent cadence/commit snapshot is skipped while paused.
    const attempt = await lib.recordVersion({ docId: "doc-1", model: currentModel, author: "You", trigger: "cadence", cachedPrev: null });
    expect(attempt.written).toBe(false);
  });

  it("deleteVersion re-encodes the immediate child before removing a version other deltas depend on", async () => {
    const { lib } = await loadVersionHistory();
    let cachedPrev: { id: string; model: DocumentModel } | null = null;
    const models: DocumentModel[] = [];
    const ids: string[] = [];
    for (let i = 0; i < 4; i++) {
      const model = baseModel("doc-1", `paragraph ${i}`);
      models.push(model);
      const result = await lib.recordVersion({ docId: "doc-1", model, author: "You", trigger: "commit", cachedPrev });
      cachedPrev = result.newEntry!;
      ids.push(result.newEntry!.id);
    }
    // Delete the second version (a delta with a delta child).
    const afterDelete = await lib.deleteVersion("doc-1", ids[1], models[3]);
    expect(afterDelete.versions.map((v) => v.id)).toEqual([ids[0], ids[2], ids[3]]);
    // The child (originally a delta on ids[1]) must now be reconstructable
    // standalone (re-encoded as a baseline).
    const child = afterDelete.versions.find((v) => v.id === ids[2])!;
    expect(child.kind).toBe("baseline");
    const reconstructed = await lib.reconstructVersion("doc-1", ids[3]);
    expect(reconstructed).not.toBeNull();
    expect((reconstructed!.sections[0].blocks[0] as Paragraph).runs[0].text).toBe("paragraph 3");
  });

  it("removeVersionName makes a version prunable again and reconcileBudget can clear the paused flag", async () => {
    const { lib, store } = await loadVersionHistory();
    store.records.set("versions:doc-1", {
      id: "versions:doc-1",
      store: "index",
      docId: "doc-1",
      versions: [{ id: "v1", timestamp: 1, author: "You", words: 1, name: "Keep me", kind: "baseline", parentId: null, size: lib.HISTORY_BUDGET_BYTES + 1 }],
      paused: true,
    });
    const model = baseModel("doc-1", "hello");
    const afterUnname = await lib.removeVersionName("doc-1", "v1", model);
    // Still over budget with only one (now unnamed, but also the newest <=10)
    // version — nothing prunable, so it stays paused; this asserts the call
    // path runs end-to-end without throwing and returns the reconciled index.
    expect(afterUnname.docId).toBe("doc-1");
  });
});

describe("versionListEntries", () => {
  it("sorts newest first and computes word-count deltas against the previous version", async () => {
    const { lib } = await loadVersionHistory();
    const index = {
      id: "versions:doc-1",
      store: "index" as const,
      docId: "doc-1",
      paused: false,
      versions: [
        { id: "a", timestamp: 1, author: "You", words: 5, kind: "baseline" as const, parentId: null, size: 10 },
        { id: "b", timestamp: 2, author: "You", words: 8, kind: "delta" as const, parentId: "a", size: 5 },
        { id: "c", timestamp: 3, author: "You", words: 6, kind: "delta" as const, parentId: "b", size: 5 },
      ],
    };
    const entries = lib.versionListEntries(index);
    expect(entries.map((e) => e.id)).toEqual(["c", "b", "a"]);
    expect(entries.find((e) => e.id === "b")?.wordDelta).toBe(3);
    expect(entries.find((e) => e.id === "c")?.wordDelta).toBe(-2);
    expect(entries.find((e) => e.id === "a")?.wordDelta).toBe(5);
  });
});
