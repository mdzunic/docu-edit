// SPEC-138 — workbook version history over the shared `versions` store:
// baseline/delta encode and decode, row windows per VERSION_ROW_BAND band
// (edits at rows 5 and 900,000 → two one-row windows), added and deleted
// sheets through `order`, meta and sourceXml deltas, blob dedupe and the
// per-reference hash cache, the cell-count and byte-budget refusals, pruning
// with re-baselining, and the list's cell deltas. The idb seam is an
// in-memory map (jsdom has no IndexedDB), as in presentationVersions.spec.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { Row, Sheet, SpreadsheetModel } from "./spreadsheet/model";

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
async function load(options: { budget?: number; store?: ReturnType<typeof fakeStore>; sha?: (text: string) => Promise<string> } = {}) {
  const store = options.store ?? fakeStore();
  vi.doMock("./idb", () => ({
    getVersionsRecord: store.getVersionsRecord,
    putVersionsRecord: store.putVersionsRecord,
    deleteVersionsRecord: store.deleteVersionsRecord,
  }));
  if (options.budget !== undefined || options.sha) {
    const { budget, sha } = options;
    vi.doMock("./core/versionChain", async (importOriginal) => {
      const actual = await importOriginal<typeof import("./core/versionChain")>();
      return { ...actual, ...(budget !== undefined ? { HISTORY_BUDGET_BYTES: budget } : {}), ...(sha ? { sha256Hex: sha } : {}) };
    });
  }
  const lib = await import("./spreadsheetVersions");
  return { lib, store };
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.doUnmock("./idb");
  vi.doUnmock("./core/versionChain");
});

function row(index: number, value: string | number = `r${index}`): Row {
  return { index, cells: [{ col: 0, value }, { col: 1, value: index }] };
}

function sheet(id: string, rows: Row[], extra: Partial<Sheet> = {}): Sheet {
  return {
    id,
    name: id.toUpperCase(),
    visibility: "visible",
    kind: "grid",
    rows,
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    conditionalFormats: [],
    validations: [],
    ...extra,
  };
}

/** An ods-shaped three-sheet workbook with a retained package, an image, an
 * envelope and a source per sheet. */
function workbook(): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "book",
    title: "Budget",
    format: "ods",
    readerVersion: 1,
    sheets: [
      sheet("s1", [row(0), row(1), row(2)], { sourceXml: "<table:table table:name=\"S1\"/>" }),
      sheet("s2", [row(0), row(5), row(40)], { sourceXml: "<table:table table:name=\"S2\"/>" }),
      sheet("s3", [row(3)], { sourceXml: "<table:table table:name=\"S3\"/>" }),
    ],
    activeSheetIndex: 1,
    dateSystem: "1900",
    opaque: [
      { name: "mimetype", data: "YXBwbGljYXRpb24vdm5k" },
      { name: "styles.xml", data: "PHN0eWxlcy8+" },
      { name: "Pictures/a.png", data: "iVBORw0KGgo=" },
    ],
    media: [{ mediaId: "Pictures/a.png", contentType: "image/png", data: "iVBORw0KGgo=" }],
    odsContentEnvelope: { prefix: "<office:document-content>", suffix: "</office:document-content>" },
    styles: [{ namedStyle: "Normal" }],
    defaultStyleId: 0,
    dxfs: [],
    names: [],
    properties: { title: "Budget" },
  };
}

/** A copy-on-write edit of one cell: a new row, rows array, sheet and sheets
 * array — everything else keeps its identity (SPEC-116's invariant). */
function editCell(model: SpreadsheetModel, sheetIndex: number, rowIndex: number, value: string): SpreadsheetModel {
  const target = model.sheets[sheetIndex]!;
  const rows = target.rows.map((r) => (r.index === rowIndex ? { ...r, cells: [{ col: 0, value }, ...r.cells.slice(1)] } : r));
  if (!rows.some((r) => r.index === rowIndex)) {
    rows.push(row(rowIndex, value));
    rows.sort((a, b) => a.index - b.index);
  }
  const sheets = model.sheets.slice();
  sheets[sheetIndex] = { ...target, rows, dirty: true };
  return { ...model, sheets };
}

type Lib = Awaited<ReturnType<typeof load>>["lib"];

function record(lib: Lib, model: SpreadsheetModel, cachedPrev: Parameters<Lib["recordWorkbookVersion"]>[0]["cachedPrev"], trigger: "commit" | "cadence" | "pre-restore" | "post-restore" = "commit") {
  return lib.recordWorkbookVersion({ docId: "book", model, author: "Ada", trigger, cachedPrev });
}

function payloadOf(store: ReturnType<typeof fakeStore>, versionId: string): Record<string, unknown> {
  return (store.records.get(`version:book:${versionId}`) as { payload: Record<string, unknown> }).payload;
}

function blobs(store: ReturnType<typeof fakeStore>) {
  return [...store.records.values()].filter((r) => (r as { store?: string }).store === "blob") as { hash: string; refCount: number; data: string }[];
}

describe("baseline and delta encode/decode", () => {
  it("the first version is a baseline carrying every sheet's meta and rows, blob refs for the sources, and reconstructs the workbook", async () => {
    const { lib, store } = await load();
    const model = workbook();
    const result = await record(lib, model, null);
    expect(result.written).toBe(true);
    const entry = result.index.versions[0]!;
    expect(entry).toMatchObject({ kind: "baseline", parentId: null, sheets: 3, cells: 14, words: 0, author: "Ada", trigger: "commit" });
    const payload = payloadOf(store, entry.id) as { kind: string; sheets: { id: string; rows: Row[]; sourceXml?: { hash: string } }[]; envelope?: unknown };
    expect(payload.kind).toBe("workbook-baseline");
    expect(payload.sheets.map((s) => s.id)).toEqual(["s1", "s2", "s3"]);
    expect(payload.sheets[1]!.rows).toEqual(model.sheets[1]!.rows);
    expect(payload.sheets[0]!.sourceXml?.hash).toMatch(/^[0-9a-f]{64}$/);
    // No bytes inline: the package, the image, the sources and the envelope are refs.
    const json = JSON.stringify(payload);
    expect(json).not.toContain("PHN0eWxlcy8+");
    expect(json).not.toContain("table:name");
    expect(json).not.toContain("office:document-content");

    expect(await lib.reconstructWorkbookVersion("book", entry.id)).toEqual(model);
  });

  it("a cell edit's delta carries that sheet's one-row window and nothing of the untouched sheets", async () => {
    const { lib, store } = await load();
    const v1 = workbook();
    const first = await record(lib, v1, null);
    const v2 = editCell(v1, 1, 5, "edited");
    const second = await record(lib, v2, first.newEntry);
    const entry = second.index.versions[1]!;
    expect(entry).toMatchObject({ kind: "delta", parentId: first.newEntry!.id, sheets: 3, cells: 14 });
    const payload = payloadOf(store, entry.id) as { order: string[]; sheets: { id: string; kind: string; rows?: unknown[]; meta?: unknown; sourceXml?: unknown }[]; fields: object; opaque?: unknown; media?: unknown; envelope?: unknown };
    expect(payload.order).toEqual(["s1", "s2", "s3"]);
    expect(payload.sheets).toHaveLength(1);
    expect(payload.sheets[0]).toMatchObject({ id: "s2", kind: "changed", rows: [{ band: 0, prefixLen: 1, suffixLen: 1, replaced: [v2.sheets[1]!.rows[1]] }] });
    // `dirty` was set: the meta goes whole; the source did not change.
    expect(payload.sheets[0]!.meta).toMatchObject({ dirty: true, name: "S2" });
    expect(payload.sheets[0]!.sourceXml).toBeUndefined();
    expect(payload.fields).toEqual({});
    expect(payload.opaque).toBeUndefined();
    expect(payload.media).toBeUndefined();
    expect(payload.envelope).toBeUndefined();
    expect(await lib.reconstructWorkbookVersion("book", entry.id)).toEqual(v2);
    expect(await lib.reconstructWorkbookVersion("book", first.newEntry!.id)).toEqual(v1);
  });

  it("a sheet whose rows array is the predecessor's contributes nothing", async () => {
    const { lib, store } = await load();
    const v1 = workbook();
    const first = await record(lib, v1, null);
    const second = await record(lib, { ...v1, sheets: v1.sheets.slice(), title: "Budget 2027" }, first.newEntry);
    const payload = payloadOf(store, second.newEntry!.id) as { sheets: unknown[]; fields: Record<string, unknown> };
    expect(payload.sheets).toEqual([]);
    expect(payload.fields).toEqual({ title: "Budget 2027" });
  });

  it("a cleared field is carried as null and reconstructs absent", async () => {
    const { lib } = await load();
    const v1 = workbook();
    const first = await record(lib, v1, null);
    const { properties: _gone, ...rest } = v1;
    const v2: SpreadsheetModel = { ...rest, sheets: v1.sheets.slice() };
    const second = await record(lib, v2, first.newEntry);
    const back = await lib.reconstructWorkbookVersion("book", second.newEntry!.id);
    expect(back && "properties" in back).toBe(false);
    expect(back).toEqual(v2);
  });

  it("builds a fresh baseline after a reload: every stored row is new, so the delta is a wholesale rewrite", async () => {
    const { lib } = await load();
    const v1 = workbook();
    await record(lib, v1, null);
    const second = await record(lib, editCell(v1, 0, 1, "after reload"), null);
    expect(second.index.versions[1]!.kind).toBe("baseline");
  });
});

describe("row windows over row identity", () => {
  it("edits at row 5 and row 900,000 cost two one-row windows, not a 900,000-row window", async () => {
    const { lib, store } = await load();
    const tall: Row[] = [];
    for (let index = 0; index < 1_000_000; index += 1000) {
      tall.push(row(index));
    }
    tall.splice(1, 0, row(5));
    const v1: SpreadsheetModel = { ...workbook(), sheets: [sheet("tall", tall)] };
    const first = await record(lib, v1, null);
    let v2 = editCell(v1, 0, 5, "top");
    v2 = editCell(v2, 0, 900_000, "bottom");
    const second = await record(lib, v2, first.newEntry);
    const payload = payloadOf(store, second.newEntry!.id) as { sheets: { rows: { band: number; prefixLen: number; suffixLen: number; replaced: Row[] }[] }[] };
    const windows = payload.sheets[0]!.rows;
    expect(windows.map((w) => [w.band, w.replaced.map((r) => r.index)])).toEqual([
      [0, [5]],
      [Math.floor(900_000 / lib.VERSION_ROW_BAND), [900_000]],
    ]);
    expect(lib.VERSION_ROW_BAND).toBe(4096);
    expect(await lib.reconstructWorkbookVersion("book", second.newEntry!.id)).toEqual(v2);
  });

  it("rowWindows and applyRowWindows round trip inserts, deletions and a new band", async () => {
    const { lib } = await load();
    const before = [row(0), row(1), row(4095), row(4096), row(9000)];
    const after = [before[0]!, row(2), before[2]!, row(5000), before[4]!, row(20_000)];
    const windows = lib.rowWindows(before, after);
    expect(windows.map((w) => w.band)).toEqual([0, 1, 4]);
    expect(lib.applyRowWindows(before, windows)).toEqual(after);
    expect(lib.rowWindows(before, before)).toEqual([]);
  });

  it("a structural edit that re-indexes every row lands in the wholesale-rewrite rule as a baseline", async () => {
    const { lib } = await load();
    const rows = Array.from({ length: 200 }, (_, i) => row(i));
    const v1: SpreadsheetModel = { ...workbook(), sheets: [sheet("grid", rows)] };
    const first = await record(lib, v1, null);
    // Insert a row at the top: every row below moves down one index.
    const shifted = [row(0, "new"), ...rows.map((r) => ({ ...r, index: r.index + 1 }))];
    const v2: SpreadsheetModel = { ...v1, sheets: [{ ...v1.sheets[0]!, rows: shifted }] };
    const second = await record(lib, v2, first.newEntry);
    expect(second.index.versions[1]!.kind).toBe("baseline");
    expect(await lib.reconstructWorkbookVersion("book", second.newEntry!.id)).toEqual(v2);
  });
});

describe("sheets added, deleted, renamed and re-sourced", () => {
  it("an added sheet goes whole, a deleted one drops out of `order`, a renamed one carries its meta", async () => {
    const { lib, store } = await load();
    const v1 = workbook();
    const first = await record(lib, v1, null);
    const added = sheet("s4", [row(0, "new sheet")]);
    const v2: SpreadsheetModel = {
      ...v1,
      sheets: [{ ...v1.sheets[0]!, name: "Renamed" }, v1.sheets[1]!, added],
      structureDirty: true,
    };
    const second = await record(lib, v2, first.newEntry);
    const payload = payloadOf(store, second.newEntry!.id) as { order: string[]; sheets: { id: string; kind: string; meta?: { name: string }; rows?: unknown }[]; fields: object };
    expect(payload.order).toEqual(["s1", "s2", "s4"]);
    expect(payload.sheets.map((s) => [s.id, s.kind])).toEqual([
      ["s1", "changed"],
      ["s4", "added"],
    ]);
    expect(payload.sheets[0]!.meta?.name).toBe("Renamed");
    expect(payload.sheets[0]!.rows).toBeUndefined();
    expect(payload.fields).toEqual({ structureDirty: true });
    expect(second.index.versions[1]).toMatchObject({ sheets: 3, cells: 14 });
    expect(await lib.reconstructWorkbookVersion("book", second.newEntry!.id)).toEqual(v2);
  });

  it("a rewritten source is carried as a new blob ref; the unchanged sources and parts keep theirs", async () => {
    const { lib, store } = await load();
    const v1 = workbook();
    const first = await record(lib, v1, null);
    const blobCount = blobs(store).length;
    const v2: SpreadsheetModel = { ...v1, sheets: v1.sheets.map((s, i) => (i === 2 ? { ...s, sourceXml: "<table:table table:name=\"S3\" moved=\"1\"/>" } : s)) };
    const second = await record(lib, v2, first.newEntry);
    const payload = payloadOf(store, second.newEntry!.id) as { sheets: { id: string; sourceXml?: { hash: string }; meta?: unknown; rows?: unknown }[] };
    expect(payload.sheets).toHaveLength(1);
    expect(payload.sheets[0]!.id).toBe("s3");
    expect(payload.sheets[0]!.sourceXml?.hash).toMatch(/^[0-9a-f]{64}$/);
    expect(payload.sheets[0]!.meta).toBeUndefined();
    expect(payload.sheets[0]!.rows).toBeUndefined();
    // Only the rewritten source added a blob.
    expect(blobs(store).length).toBe(blobCount + 1);
    expect((await lib.reconstructWorkbookVersion("book", second.newEntry!.id))?.sheets[2]!.sourceXml).toContain('moved="1"');
  });

  it("restores the live workbook's own package, source strings and envelope when their content is unchanged", async () => {
    const { lib } = await load();
    const live = workbook();
    const first = await record(lib, live, null);
    const back = await lib.reconstructWorkbookVersion("book", first.newEntry!.id, live);
    expect(back?.opaque).toBe(live.opaque);
    expect(back?.media).toBe(live.media);
    expect(back?.odsContentEnvelope).toBe(live.odsContentEnvelope);
    expect(back?.sheets[0]!.sourceXml).toBe(live.sheets[0]!.sourceXml);
  });

  it("is null when a blob is missing", async () => {
    const { lib, store } = await load();
    const first = await record(lib, workbook(), null);
    store.records.delete(`blob:book:${blobs(store)[0]!.hash}`);
    expect(await lib.reconstructWorkbookVersion("book", first.newEntry!.id)).toBeNull();
  });
});

describe("blobs", () => {
  it("dedupes a media part that duplicates a package part, counting each blob once per payload", async () => {
    const { lib, store } = await load();
    await record(lib, workbook(), null);
    // Three package parts (the image shared with media), three sources, one envelope.
    expect(blobs(store)).toHaveLength(7);
    expect(blobs(store).every((blob) => blob.refCount === 1)).toBe(true);
  });

  it("hashes each source at most once per reference — a 30 MB table's source costs one hash, not one per version", async () => {
    let hashed = 0;
    const { lib } = await load({
      sha: async (text: string) => {
        hashed += text.length > 1000 ? 1 : 0;
        const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
        return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, "0")).join("");
      },
    });
    const v1 = workbook();
    v1.sheets[0] = { ...v1.sheets[0]!, sourceXml: `<table:table>${"<table:table-row/>".repeat(200)}</table:table>` };
    let prev = (await record(lib, v1, null)).newEntry;
    let model = v1;
    for (let i = 0; i < 3; i++) {
      model = editCell(model, 0, 1, `edit ${i}`);
      prev = (await record(lib, model, prev)).newEntry;
    }
    expect(hashed).toBe(1);
  });
});

describe("the over-budget degrade", () => {
  it("refuses a workbook over WORKBOOK_HISTORY_MAX_CELLS before any encode — the index latch is the only write", async () => {
    const { lib, store } = await load();
    expect(lib.WORKBOOK_HISTORY_MAX_CELLS).toBe(250_000);
    const rows = Array.from({ length: 125_001 }, (_, i) => row(i));
    const huge: SpreadsheetModel = { ...workbook(), sheets: [sheet("big", rows)] };
    const stringify = vi.spyOn(JSON, "stringify");
    const result = await record(lib, huge, null);
    expect(result.written).toBe(false);
    expect(result.index).toMatchObject({ paused: true, pausedReason: "too-large", versions: [] });
    expect(stringify.mock.calls.some(([value]) => value === rows || (value as { sheets?: unknown })?.sheets !== undefined)).toBe(false);
    stringify.mockRestore();
    expect(store.putVersionsRecord.mock.calls.map(([r]) => (r as unknown as { store: string }).store)).toEqual(["index"]);
    expect(lib.WORKBOOK_HISTORY_TOO_LARGE_MESSAGE).toBe("Version history off — this workbook is larger than the 50 MB history budget");
  });

  it("refuses a smaller workbook whose blobs plus baseline exceed the budget, before any write, and stays latched", async () => {
    const { lib, store } = await load({ budget: 300 });
    const result = await record(lib, workbook(), null);
    expect(result.written).toBe(false);
    expect(result.index).toMatchObject({ paused: true, pausedReason: "too-large" });
    expect(store.putVersionsRecord.mock.calls.map(([r]) => (r as unknown as { store: string }).store)).toEqual(["index"]);
    store.putVersionsRecord.mockClear();
    expect((await record(lib, workbook(), null, "pre-restore")).written).toBe(false);
    expect(store.putVersionsRecord).not.toHaveBeenCalled();
  });
});

describe("pruning", () => {
  async function seed(lib: Lib, count: number) {
    let model: SpreadsheetModel = { ...workbook(), sheets: [sheet("grid", Array.from({ length: 60 }, (_, i) => row(i)))] };
    let cachedPrev = null;
    const ids: string[] = [];
    const models: SpreadsheetModel[] = [];
    for (let i = 0; i < count; i++) {
      if (i > 0) {
        model = editCell(model, 0, i, `edit ${i}`);
      }
      const result = await record(lib, model, cachedPrev);
      cachedPrev = result.newEntry;
      ids.push(result.newEntry!.id);
      models.push(model);
    }
    return { ids, models };
  }

  function used(store: ReturnType<typeof fakeStore>, index: { versions: { size: number }[] }) {
    return (
      index.versions.reduce((sum, v) => sum + v.size, 0) +
      blobs(store).reduce((sum, b) => sum + (b as unknown as { size: number }).size, 0)
    );
  }

  it("prunes the oldest unretained versions, re-baselining the oldest retained dependant", async () => {
    const seeded = await load();
    const { ids, models } = await seed(seeded.lib, 12);
    const budget = used(seeded.store, await seeded.lib.loadWorkbookIndex("book")) - 1;
    vi.resetModules();
    const { lib } = await load({ budget, store: seeded.store });
    const after = await lib.reconcileWorkbookBudget("book", models[11]!);
    expect(after.paused).toBe(false);
    expect(after.versions.map((v) => v.id)).toEqual(ids.slice(2));
    expect(after.versions[0]).toMatchObject({ kind: "baseline", parentId: null });
    for (const [i, id] of ids.entries()) {
      if (i >= 2) {
        expect(await lib.reconstructWorkbookVersion("book", id)).toEqual(models[i]);
      }
    }
  });

  it("deleteWorkbookVersion re-encodes the immediate child; names come and go", async () => {
    const { lib } = await load();
    const { ids, models } = await seed(lib, 4);
    const after = await lib.deleteWorkbookVersion("book", ids[1]!, models[3]!);
    expect(after.versions.map((v) => v.id)).toEqual([ids[0], ids[2], ids[3]]);
    expect(after.versions[1]!.kind).toBe("baseline");
    expect(await lib.reconstructWorkbookVersion("book", ids[3]!)).toEqual(models[3]);
    expect((await lib.nameWorkbookVersion("book", ids[0]!, "  Before review ")).versions[0]!.name).toBe("Before review");
    expect((await lib.removeWorkbookVersionName("book", ids[0]!, models[3]!)).versions[0]!.name).toBeUndefined();
  });
});

describe("workbookVersionEntries", () => {
  it("lists newest first with each version's cell-count delta", async () => {
    const { lib } = await load();
    const entries = lib.workbookVersionEntries({
      id: "versions:book",
      store: "index",
      docId: "book",
      paused: false,
      versions: [
        { id: "a", timestamp: 1, author: "Ada", words: 0, sheets: 2, cells: 40, kind: "baseline", parentId: null, size: 10 },
        { id: "b", timestamp: 2, author: "Ada", words: 0, sheets: 2, cells: 52, kind: "delta", parentId: "a", size: 5 },
        { id: "c", timestamp: 3, author: "Ada", words: 0, sheets: 3, cells: 49, kind: "delta", parentId: "b", size: 5 },
        { id: "d", timestamp: 4, author: "Ada", words: 0, sheets: 3, cells: 49, kind: "delta", parentId: "c", size: 5 },
      ],
    });
    expect(entries.map((e) => [e.id, e.cellDelta])).toEqual([
      ["d", 0],
      ["c", -3],
      ["b", 12],
      ["a", 40],
    ]);
  });
});
