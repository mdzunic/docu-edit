// SPEC-137 "The storage layout" in the store: a workbook persists as one
// record per 512-row band of row indices per sheet, one source record and a
// slim shell carrying the `sheetStore` manifest; a persist rewrites only the
// chunks whose slice changed by identity, deletes emptied chunks and a
// deleted sheet's chunks, rewrites the source record only when a source
// changed, and commits all of it with the shell in ONE writeBatch. A
// workbook is resident only while needed (getSpreadsheetShell /
// getSpreadsheet / loadSpreadsheet / releaseSpreadsheet), SPEC-116-era
// records migrate lazily (on load, and in an idle pass), a manifest naming a
// record storage does not return opens read-only and is never written back,
// and the landing page reads no chunk or source record (asserted with the
// real lib/idb.ts over the in-memory IndexedDB).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mount } from "@vue/test-utils";
import type { WriteBatchOp } from "./core/idb";
import type { Row, SpreadsheetModel } from "./spreadsheet/model";
import { FakeIndexedDb } from "./core/fakeIndexedDb";

/** An `idb` seam over a Map: writeBatch applies its ops, getRecord serves
 * them back, every call recorded. */
function makeIdb(initial: Array<{ id: string }> = []) {
  const data = new Map<string, unknown>(initial.map((record) => [record.id, structuredClone(record)]));
  const batches: WriteBatchOp[][] = [];
  const idb = {
    data,
    batches,
    loadAllRecords: vi.fn(async () =>
      [...data.entries()].filter(([key]) => !key.startsWith("sheetrows:") && !key.startsWith("sheetsource:")).map(([, value]) => structuredClone(value)),
    ),
    putRecord: vi.fn(async (record: { id: string }) => {
      data.set(record.id, structuredClone(record));
    }),
    deleteRecord: vi.fn(async (key: string) => {
      data.delete(key);
    }),
    getRecord: vi.fn(async (key: string) => (data.has(key) ? structuredClone(data.get(key)) : undefined)),
    writeBatch: vi.fn(async (ops: WriteBatchOp[]) => {
      batches.push(ops);
      for (const op of ops) {
        if ("put" in op) {
          const record = op.put as { id: string };
          data.set(record.id, structuredClone(record));
        } else if ("delete" in op) {
          data.delete(op.delete);
        } else {
          for (const key of [...data.keys()]) {
            if (key >= op.deleteRange[0] && key <= op.deleteRange[1]) {
              data.delete(key);
            }
          }
        }
      }
    }),
  };
  return idb;
}

async function loadStore(initial: Array<{ id: string }> = []) {
  const idb = makeIdb(initial);
  vi.doMock("./idb", () => idb);
  const store = await import("./documentStore");
  const { createBlankWorkbook } = await import("./spreadsheet/sheets");
  return { store, idb, createBlankWorkbook };
}

/** Real ticks: the write-through yields a macrotask before each batch. */
async function flush(ticks = 12) {
  for (let i = 0; i < ticks; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

function rowsOf(count: number, step = 1, cols = 2): Row[] {
  return Array.from({ length: count }, (_, i) => ({
    index: i * step,
    cells: Array.from({ length: cols }, (_, col) => ({ col, value: col === 0 ? i : `r${i}c${col}` })),
  }));
}

/** A two-sheet workbook: "Data" with `dataRows`, "Notes" with three rows,
 * a retained package and an ods-style source on Notes. */
function workbook(createBlankWorkbook: () => SpreadsheetModel, dataRows: Row[]): SpreadsheetModel {
  const model = createBlankWorkbook();
  const first = model.sheets[0]!;
  model.sheets = [
    { ...first, id: "data", name: "Data", rows: dataRows },
    { ...first, id: "notes", name: "Notes", rows: rowsOf(3), sourceXml: "<table:table/>" },
  ];
  model.opaque = [{ name: "xl/workbook.xml", data: "PHdvcmtib29rLz4=" }];
  return model;
}

function puts(ops: WriteBatchOp[]): Array<{ id: string; [key: string]: unknown }> {
  return ops.flatMap((op) => ("put" in op ? [op.put as { id: string }] : []));
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.doUnmock("./idb");
});

describe("chunk keys and bands", () => {
  it("writes one record per 512-row band of row indices, a source record and a slim shell with the manifest — in one batch", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    expect(store.SHEET_ROWS_CHUNK).toBe(512);
    expect(store.SHEET_SOURCE_KEY_PREFIX).toBe("sheetsource:");
    const sparse: Row[] = [0, 511, 512, 2000].map((index) => ({ index, cells: [{ col: 0, value: index }] }));
    const model = workbook(createBlankWorkbook, sparse);
    store.storeSpreadsheet(model);
    await flush();

    expect(idb.writeBatch).toHaveBeenCalledTimes(1);
    const written = puts(idb.batches[0]!);
    const chunks = written.filter((record) => record.id.startsWith("sheetrows:"));
    expect(chunks.map((record) => record.id)).toEqual([
      `sheetrows:${model.id}:data:0000`,
      `sheetrows:${model.id}:data:0001`,
      `sheetrows:${model.id}:data:0003`,
      `sheetrows:${model.id}:notes:0000`,
    ]);
    expect(chunks[0]).toMatchObject({ docId: model.id, sheetId: "data", chunk: 0 });
    expect((chunks[0]!.rows as Row[]).map((row) => row.index)).toEqual([0, 511]);
    expect((chunks[1]!.rows as Row[]).map((row) => row.index)).toEqual([512]);
    const source = written.find((record) => record.id === `sheetsource:${model.id}`)!;
    expect(source).toMatchObject({ docId: model.id, opaque: model.opaque, sheetSources: { notes: "<table:table/>" } });
    const shell = written.at(-1) as unknown as { id: string; model: SpreadsheetModel; sheetStore: { chunks: Record<string, number[]>; storedBytes: number } };
    expect(shell.id).toBe(model.id);
    expect(shell.model.opaque).toEqual([]);
    expect(shell.model.sheets.map((sheet) => sheet.rows)).toEqual([[], []]);
    expect(shell.model.sheets.some((sheet) => "sourceXml" in sheet)).toBe(false);
    expect(shell.sheetStore.chunks).toEqual({ data: [0, 1, 3], notes: [0] });
    // The in-memory model is untouched and resident.
    expect(store.getSpreadsheet(model.id)).toBe(model);
    expect(model.sheets[0]!.rows).toBe(sparse);
  });
});

describe("a persist rewrites only what changed", () => {
  async function stored25k() {
    const loaded = await loadStore();
    const model = workbook(loaded.createBlankWorkbook, rowsOf(25_000));
    loaded.store.storeSpreadsheet(model);
    await flush();
    loaded.idb.writeBatch.mockClear();
    loaded.idb.batches.length = 0;
    return { ...loaded, model };
  }

  it("one edited cell of a 25,000-row sheet writes exactly one chunk record and the shell, in one writeBatch", async () => {
    const { store, idb, model } = await stored25k();
    const data = model.sheets[0]!;
    const rows = [...data.rows];
    rows[20_000] = { ...rows[20_000]!, cells: [{ col: 0, value: "edited" }, ...rows[20_000]!.cells.slice(1)] };
    model.sheets = [{ ...data, rows }, model.sheets[1]!];
    await store.storeSpreadsheetEdit(model);

    expect(idb.writeBatch).toHaveBeenCalledTimes(1);
    const ops = idb.batches[0]!;
    expect(ops.map((op) => ("put" in op ? (op.put as { id: string }).id : JSON.stringify(op)))).toEqual([
      `sheetrows:${model.id}:data:0039`,
      model.id,
    ]);
    expect((puts(ops)[0]!.rows as Row[])[20_000 - 39 * 512]!.cells[0]!.value).toBe("edited");
  });

  it("a rename or a star writes the shell alone", async () => {
    const { store, idb, model } = await stored25k();
    store.renameDocument(model.id, "Renamed");
    await flush();
    store.setStarred(model.id, true);
    await flush();
    expect(idb.batches.map((ops) => ops.map((op) => ("put" in op ? (op.put as { id: string }).id : "op")))).toEqual([[model.id], [model.id]]);
  });

  it("deletes a chunk that became empty and drops a deleted sheet's chunks", async () => {
    const { store, idb, model } = await stored25k();
    const data = model.sheets[0]!;
    // Rows 1,024..1,535 (chunk 2) removed; Notes deleted.
    model.sheets = [{ ...data, rows: data.rows.filter((row) => row.index < 1024 || row.index >= 1536) }];
    await store.storeSpreadsheetEdit(model);

    expect(idb.writeBatch).toHaveBeenCalledTimes(1);
    const ops = idb.batches[0]!;
    expect(ops).toContainEqual({ delete: `sheetrows:${model.id}:data:0002` });
    expect(ops).toContainEqual({ deleteRange: [`sheetrows:${model.id}:notes:`, `sheetrows:${model.id}:notes:￿`] });
    // Notes' sourceXml left the workbook: the source record is rewritten.
    expect(puts(ops).map((record) => record.id)).toEqual([`sheetsource:${model.id}`, model.id]);
    const shell = puts(ops).at(-1) as unknown as { sheetStore: { chunks: Record<string, number[]> } };
    expect(shell.sheetStore.chunks.data).not.toContain(2);
    expect(Object.keys(shell.sheetStore.chunks)).toEqual(["data"]);
    expect([...idb.data.keys()].filter((key) => key.includes(":notes:"))).toEqual([]);
  });

  it("estimates storedBytes at persist time and libraryStorageUsage reports it without serialising the workbook", async () => {
    const { store, idb, model } = await stored25k();
    const shell = idb.data.get(model.id) as { sheetStore: { storedBytes: number } };
    const { workbookStoredBytes } = await import("./workbookStorage");
    expect(shell.sheetStore.storedBytes).toBe(workbookStoredBytes(model));
    expect(shell.sheetStore.storedBytes).toBeGreaterThan(25_000 * 2 * 16);
    const stringify = vi.spyOn(JSON, "stringify");
    const usage = store.libraryStorageUsage();
    expect(usage.byDocument).toEqual([{ id: model.id, bytes: shell.sheetStore.storedBytes }]);
    expect(stringify.mock.calls.some(([value]) => (value as { kind?: string } | null)?.kind === "spreadsheet")).toBe(false);
  });
});

describe("residency", () => {
  it("release drops rows and sources once persistence flushed; load rejoins them with one get per chunk; concurrent loads share one promise", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(1_200));
    store.storeSpreadsheet(model);
    // Released before the idle write ran: the release flushes it first.
    await store.releaseSpreadsheet(model.id);
    expect(idb.writeBatch).toHaveBeenCalled();
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
    const shell = store.getSpreadsheetShell(model.id)!;
    expect(shell.sheets.map((sheet) => sheet.rows)).toEqual([[], []]);
    expect(shell.opaque).toEqual([]);
    expect(store.listDocuments().map((entry) => entry.id)).toEqual([model.id]);

    idb.getRecord.mockClear();
    const first = store.loadSpreadsheet(model.id);
    const second = store.loadSpreadsheet(model.id);
    expect(second).toBe(first);
    const loaded = (await first)!;
    expect(idb.getRecord.mock.calls.map(([key]) => key)).toEqual([
      `sheetsource:${model.id}`,
      `sheetrows:${model.id}:data:0000`,
      `sheetrows:${model.id}:data:0001`,
      `sheetrows:${model.id}:data:0002`,
      `sheetrows:${model.id}:notes:0000`,
    ]);
    expect(loaded.sheets.map((sheet) => sheet.rows)).toEqual(model.sheets.map((sheet) => sheet.rows));
    expect(loaded.sheets[1]!.sourceXml).toBe("<table:table/>");
    expect(loaded.opaque).toEqual(model.opaque);
    expect(store.getSpreadsheet(model.id)).toBe(loaded);
    expect(await store.loadSpreadsheet(model.id)).toBe(loaded);

    // The latch was seeded from storage: a rename writes the shell alone.
    idb.writeBatch.mockClear();
    store.renameDocument(model.id, "Again");
    await flush();
    expect(idb.batches.at(-1)!.map((op) => ("put" in op ? (op.put as { id: string }).id : "op"))).toEqual([model.id]);
  });

  it("release waits for pending persistence, and the store never writes rows or sources for a workbook that is not resident", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(600));
    store.storeSpreadsheet(model);
    await flush();
    const rows = [...model.sheets[0]!.rows];
    rows[0] = { index: 0, cells: [{ col: 0, value: "late" }] };
    model.sheets = [{ ...model.sheets[0]!, rows }, model.sheets[1]!];
    void store.storeSpreadsheetEdit(model);
    await store.releaseSpreadsheet(model.id);
    expect((idb.data.get(`sheetrows:${model.id}:data:0000`) as { rows: Row[] }).rows[0]!.cells[0]!.value).toBe("late");
    expect(store.getSpreadsheet(model.id)).toBeUndefined();

    idb.writeBatch.mockClear();
    idb.batches.length = 0;
    store.setStarred(model.id, true);
    await flush();
    expect(idb.batches).toHaveLength(1);
    expect(idb.batches[0]!.map((op) => ("put" in op ? (op.put as { id: string }).id : "op"))).toEqual([model.id]);
    expect((idb.data.get(`sheetrows:${model.id}:data:0000`) as { rows: Row[] }).rows).toHaveLength(512);
  });

  it("keeps a workbook resident when storage could not take it", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    idb.writeBatch.mockRejectedValue(new Error("quota"));
    const model = workbook(createBlankWorkbook, rowsOf(10));
    store.storeSpreadsheet(model);
    await store.releaseSpreadsheet(model.id);
    expect(store.getSpreadsheet(model.id)).toBe(model);
    expect(store.isStorageAvailable()).toBe(false);
  });

  it("the library's duplicate loads a workbook that is not resident, copies it whole and releases both", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(700));
    store.storeSpreadsheet(model);
    await store.releaseSpreadsheet(model.id);
    expect(store.duplicateDocument(model.id)).toBeUndefined();
    const copy = (await store.duplicateStoredDocument(model.id))!;
    await flush(20);
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
    expect(store.getSpreadsheet(copy.id)).toBeUndefined();
    expect((idb.data.get(`sheetrows:${copy.id}:data:0001`) as { rows: Row[] }).rows).toHaveLength(188);
    expect((await store.loadSpreadsheet(copy.id))!.sheets[0]!.rows).toHaveLength(700);
  });

  it("deleteDocument removes the shell, every rows record and the source record in one batch", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(1_100));
    store.storeSpreadsheet(model);
    await flush();
    store.deleteDocument(model.id);
    await flush();
    expect(idb.batches.at(-1)).toEqual([
      { delete: model.id },
      { deleteRange: [`sheetrows:${model.id}:`, `sheetrows:${model.id}:￿`] },
      { delete: `sheetsource:${model.id}` },
    ]);
    expect([...idb.data.keys()].filter((key) => key.includes(model.id))).toEqual([]);
  });
});

describe("incomplete storage", () => {
  it("a manifest naming a chunk storage does not return opens read-only, and the partial model is never written back", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(1_100));
    store.storeSpreadsheet(model);
    await store.releaseSpreadsheet(model.id);
    idb.data.delete(`sheetrows:${model.id}:data:0001`);
    const { WORKBOOK_STORAGE_INCOMPLETE_MESSAGE } = await import("./spreadsheet/errors");
    expect(WORKBOOK_STORAGE_INCOMPLETE_MESSAGE).toBe(
      "Part of this workbook couldn't be loaded from browser storage — download it or reopen the original file",
    );

    const partial = (await store.loadSpreadsheet(model.id))!;
    expect(store.isSpreadsheetStorageIncomplete(model.id)).toBe(true);
    expect(partial.sheets[0]!.rows).toHaveLength(1_100 - 512);

    idb.writeBatch.mockClear();
    const rows = [{ index: 0, cells: [{ col: 0, value: "x" }] }, ...partial.sheets[0]!.rows.slice(1)];
    partial.sheets = [{ ...partial.sheets[0]!, rows }, partial.sheets[1]!];
    await store.storeSpreadsheetEdit(partial);
    await store.touchDocument(model.id);
    await flush();
    expect(idb.writeBatch).not.toHaveBeenCalled();
    expect(store.duplicateDocument(model.id)).toBeUndefined();
    expect((idb.data.get(`sheetrows:${model.id}:data:0000`) as { rows: Row[] }).rows[0]!.cells[0]!.value).toBe(0);
  });

  it("a missing source record is incomplete too", async () => {
    const { store, idb, createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(10));
    store.storeSpreadsheet(model);
    await store.releaseSpreadsheet(model.id);
    idb.data.delete(`sheetsource:${model.id}`);
    await store.loadSpreadsheet(model.id);
    expect(store.isSpreadsheetStorageIncomplete(model.id)).toBe(true);
  });
});

describe("SPEC-116-era records", () => {
  /** A 116-era workbook in storage: a shell with no manifest (sources
   * inline, rows emptied) and one `sheetrows:<docId>:<sheetId>` per sheet. */
  async function legacy() {
    const { createBlankWorkbook } = await loadStore();
    const model = workbook(createBlankWorkbook, rowsOf(1_000));
    vi.resetModules();
    const shell = { id: model.id, format: "xlsx", lastModified: 500, model: { ...model, sheets: model.sheets.map((sheet) => ({ ...sheet, rows: [] })) } };
    const perSheet = model.sheets.map((sheet) => ({ id: `sheetrows:${model.id}:${sheet.id}`, docId: model.id, sheetId: sheet.id, rows: sheet.rows }));
    return { model, records: [shell, ...perSheet] };
  }

  it("loadSpreadsheet rejoins the per-sheet records, then rewrites chunks, a source record and a slim shell and deletes the legacy records — in one writeBatch", async () => {
    const { model, records } = await legacy();
    const { store, idb } = await loadStore(records);
    await store.initLibrary();
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
    const loaded = (await store.loadSpreadsheet(model.id))!;
    expect(loaded.sheets.map((sheet) => sheet.rows)).toEqual(model.sheets.map((sheet) => sheet.rows));
    await flush();
    const migration = idb.batches[0]!;
    expect(migration[0]).toEqual({ deleteRange: [`sheetrows:${model.id}:`, `sheetrows:${model.id}:￿`] });
    expect(puts(migration).map((record) => record.id)).toEqual([
      `sheetrows:${model.id}:data:0000`,
      `sheetrows:${model.id}:data:0001`,
      `sheetrows:${model.id}:notes:0000`,
      `sheetsource:${model.id}`,
      model.id,
    ]);
    expect(idb.data.has(`sheetrows:${model.id}:data`)).toBe(false);
    const shell = idb.data.get(model.id) as { model: SpreadsheetModel; sheetStore?: unknown };
    expect(shell.sheetStore).toBeDefined();
    expect(shell.model.opaque).toEqual([]);
  });

  it("an idle pass after initLibrary migrates the workbooks nobody opens and leaves them released", async () => {
    const { model, records } = await legacy();
    const { store, idb } = await loadStore(records);
    await store.initLibrary();
    await vi.waitFor(() => expect((idb.data.get(model.id) as { sheetStore?: unknown }).sheetStore).toBeDefined());
    await flush();
    expect(idb.writeBatch).toHaveBeenCalledTimes(1);
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
    expect(idb.data.has(`sheetrows:${model.id}:data:0001`)).toBe(true);
  });
});

describe("the landing page reads no cell (real lib/idb.ts over an in-memory IndexedDB)", () => {
  let fake: FakeIndexedDb;
  beforeEach(() => {
    fake = new FakeIndexedDb();
    fake.install();
  });
  afterEach(() => {
    fake.uninstall();
  });

  it("20 stored workbooks list, sort, search and render their cards from shells and preview alone", async () => {
    const idbModule = await import("./idb");
    let store = await import("./documentStore");
    expect(idbModule.LAZY_RECORD_KEY_PREFIXES).toEqual([store.SHEET_ROWS_KEY_PREFIX, store.SHEET_SOURCE_KEY_PREFIX]);
    const { createBlankWorkbook } = await import("./spreadsheet/sheets");
    for (let i = 0; i < 20; i++) {
      const model = workbook(createBlankWorkbook, rowsOf(1_100));
      model.title = `Library ${String(i).padStart(2, "0")}`;
      store.storeSpreadsheet(model);
      await store.releaseSpreadsheet(model.id);
    }
    const keys = [...fake.stores.get("documents")!.keys()];
    expect(keys.filter((key) => key.startsWith("sheetrows:"))).toHaveLength(20 * 4);
    expect(keys.filter((key) => key.startsWith("sheetsource:"))).toHaveLength(20);

    // A fresh session over the same storage.
    vi.resetModules();
    fake.log.length = 0;
    store = await import("./documentStore");
    await store.initLibrary();
    const { useLibrary } = await import("../composables/useLibrary");
    const lib = useLibrary();
    lib.refresh();
    expect(lib.visible.value).toHaveLength(20);
    lib.setSortKey("title");
    expect(lib.visible.value[0]!.title).toBe("Library 00");
    lib.setSortKey("size");
    lib.setSearch("lib 07");
    expect(lib.visible.value.map((entry) => entry.title)).toEqual(["Library 07"]);
    const { default: LibraryCard } = await import("../components/library/LibraryCard.vue");
    const entry = store.listDocuments()[0]!;
    const card = mount(LibraryCard, { props: { entry, now: Date.now() } });
    await vi.waitFor(() => expect(card.find('[data-testid="sheet-miniature"]').exists()).toBe(true));
    await vi.waitFor(() => expect(card.get('[data-testid="row-counts"]').text()).toBe("2 sheets"));
    card.unmount();

    const read = fake.readKeys();
    expect(read.length).toBeGreaterThanOrEqual(20);
    expect(read.filter((key) => key.startsWith("sheetrows:") || key.startsWith("sheetsource:"))).toEqual([]);
    expect(store.getSpreadsheet(entry.id)).toBeUndefined();
  });
});
