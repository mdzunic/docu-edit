// SPEC-116 in the store: a workbook is the fourth kind. Since SPEC-137 its
// rows persist as 512-row chunk records (`sheetrows:<docId>:<sheetId>:<chunk>`),
// its sources as one `sheetsource:<docId>` record and the shell slim — written
// only where a reference changed (the model's copy-on-write invariant), so a
// rename, a star or a library listing never structured-clones a cell; rejoined
// by loadSpreadsheet (lazy residency replaced SPEC-116's "every read path
// rejoins"), deleted with the document (and with a sheet that no longer
// exists), counted by libraryStorageUsage from the manifest. A record stored
// by an older reader is re-read from its retained package on load, keeping
// its id, title, starred state and sheet ids, and written back once.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { WriteBatchOp } from "./core/idb";
import { buildMiniBookXlsxBytes, buildSpecMacrosXlsmBytes } from "./spreadsheet/fixtures/build";
import type { Row, SpreadsheetModel } from "./spreadsheet/model";
import { SPREADSHEET_READER_VERSION } from "./spreadsheet/readerVersion";

/** The idb seam over a Map: writeBatch applies its ops, getRecord serves them. */
function makeIdbMock(loaded: Array<{ id: string }> = []) {
  const data = new Map<string, unknown>(loaded.map((record) => [record.id, structuredClone(record)]));
  const batches: WriteBatchOp[][] = [];
  return {
    data,
    batches,
    loadAllRecords: vi.fn(async () =>
      [...data.entries()].filter(([key]) => !key.startsWith("sheetrows:") && !key.startsWith("sheetsource:")).map(([, value]) => structuredClone(value)),
    ),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    getRecord: vi.fn(async (key: string) => (data.has(key) ? structuredClone(data.get(key)) : undefined)),
    writeBatch: vi.fn(async (ops: WriteBatchOp[]) => {
      batches.push(ops);
      for (const op of ops) {
        if ("put" in op) {
          data.set((op.put as { id: string }).id, structuredClone(op.put));
        } else if ("delete" in op) {
          data.delete(op.delete);
        } else {
          for (const key of data.keys()) {
            if (key >= op.deleteRange[0] && key <= op.deleteRange[1]) {
              data.delete(key);
            }
          }
        }
      }
    }),
  };
}

async function loadStore(loaded: Array<{ id: string }> = []) {
  const idb = makeIdbMock(loaded);
  vi.doMock("./idb", () => idb);
  const store = await import("./documentStore");
  const { readSpreadsheetBytes } = await import("./spreadsheet/open");
  const workbook = async (): Promise<SpreadsheetModel> =>
    readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
  return { store, idb, workbook, readSpreadsheetBytes };
}

// The write-through is deferred to idle and yields a macrotask before each
// batch (see documentStore.library.spec.ts): several real ticks catch up.
async function flushWrites() {
  for (let i = 0; i < 10; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

interface WrittenRecord {
  id: string;
  docId?: string;
  sheetId?: string;
  rows?: Row[];
  model?: SpreadsheetModel;
}

/** Every record a batch put, in order. */
function written(idb: ReturnType<typeof makeIdbMock>): WrittenRecord[] {
  return idb.batches.flatMap((ops) => ops.flatMap((op) => ("put" in op ? [op.put as WrittenRecord] : [])));
}

function rowsWrites(idb: ReturnType<typeof makeIdbMock>): WrittenRecord[] {
  return written(idb).filter((record) => record.id.startsWith("sheetrows:"));
}

function clearWrites(idb: ReturnType<typeof makeIdbMock>): void {
  idb.batches.length = 0;
  idb.writeBatch.mockClear();
}

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.doUnmock("./idb");
});

describe("the fourth kind", () => {
  it("narrows, lists and routes a workbook as a spreadsheet — and never as another kind", async () => {
    const { store, workbook, readSpreadsheetBytes } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    expect(store.isSpreadsheetModel(model)).toBe(true);
    expect(store.getSpreadsheet(model.id)).toBe(model);
    expect(store.getDocument(model.id)).toBeUndefined();
    expect(store.getPresentation(model.id)).toBeUndefined();
    expect(store.getMedia(model.id)).toBeUndefined();
    expect(store.documentKindOf(model.id)).toBe("spreadsheet");
    expect(store.listDocuments()[0]).toMatchObject({ id: model.id, format: "xlsx", kind: "spreadsheet", title: "mini-book" });
    expect(store.listDocuments()[0]!.packageVariant).toBeUndefined();

    const macros = await readSpreadsheetBytes(await buildSpecMacrosXlsmBytes(), "xlsx", "Macros");
    store.storeSpreadsheet(macros);
    expect(store.listDocuments().find((entry) => entry.id === macros.id)?.packageVariant).toBe("macroEnabled");
  });

  it("getOrCreateDocument, saveDraft and replaceDocumentModel throw on a workbook id", async () => {
    const { store, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    expect(() => store.getOrCreateDocument(model.id)).toThrow(/spreadsheet/);
    await expect(store.saveDraft(model.id, null, 0)).rejects.toThrow(/spreadsheet/);
    expect(() => store.replaceDocumentModel(model.id, { id: model.id, title: "x", sections: [], defaultStyles: {} } as never)).toThrow(
      /spreadsheet/,
    );
    expect(store.getSpreadsheet(model.id)).toBe(model);
  });

  it("storeOpenedFile dispatches a workbook to the spreadsheet store", async () => {
    const { store, workbook } = await loadStore();
    const model = await workbook();
    store.storeOpenedFile({ kind: "spreadsheet", model });
    expect(store.getSpreadsheet(model.id)).toBe(model);
  });
});

describe("the chunked row split (SPEC-137 replaces SPEC-116's per-sheet record)", () => {
  it("writes each sheet's rows to its chunk records and the shell with every sheet's rows emptied", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();

    const shell = written(idb).find((record) => record.id === model.id)!;
    expect(shell.model!.sheets.map((sheet) => sheet.rows)).toEqual([[], [], []]);
    expect(shell.model!.sheets.map((sheet) => sheet.id)).toEqual(model.sheets.map((sheet) => sheet.id));
    // The card renders from the shell alone.
    expect(shell.model!.preview?.cells[1]?.[1]).toBe("Revenue");
    const rows = rowsWrites(idb);
    expect(rows.map((record) => record.id).toSorted()).toEqual(
      model.sheets.map((sheet) => `sheetrows:${model.id}:${sheet.id}:0000`).toSorted(),
    );
    for (const sheet of model.sheets) {
      const record = rows.find((entry) => entry.sheetId === sheet.id)!;
      expect(record.docId).toBe(model.id);
      expect(record.rows).toEqual(sheet.rows);
    }
    // The in-memory model keeps its rows.
    expect(store.getSpreadsheet(model.id)!.sheets[1]!.rows).toHaveLength(500);
  });

  it("a rename, a star or a library listing writes no rows record", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    clearWrites(idb);

    store.renameDocument(model.id, "Budget");
    await flushWrites();
    store.setStarred(model.id, true);
    store.listDocuments();
    store.libraryStorageUsage();
    await flushWrites();

    expect(rowsWrites(idb)).toEqual([]);
    expect(written(idb).map((record) => record.id)).toEqual([model.id, model.id]);
    expect(written(idb)[0]!.model!.title).toBe("Budget");
  });

  it("editing one sheet's rows array writes exactly that sheet's chunk, and refreshes the preview", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    clearWrites(idb);

    // Copy-on-write: a new A1 cell, a new first row, a new rows array; every
    // other row keeps its identity.
    const summary = model.sheets[0]!;
    const [first, ...rest] = summary.rows;
    summary.rows = [{ ...first!, cells: [{ col: 0, value: "Edited" }, ...first!.cells.slice(1)] }, ...rest];
    await store.touchDocument(model.id);

    const rows = rowsWrites(idb);
    expect(rows.map((record) => record.id)).toEqual([`sheetrows:${model.id}:${summary.id}:0000`]);
    expect(rows[0]!.rows).toEqual(summary.rows);
    const shell = written(idb).find((record) => record.id === model.id)!;
    expect(shell.model!.preview?.cells[0]?.[0]).toBe("Edited");
  });

  it("deletes a removed sheet's chunks on the next persist", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    const removed = model.sheets[2]!;
    model.sheets = model.sheets.slice(0, 2);
    await store.touchDocument(model.id);
    expect(idb.batches.at(-1)).toContainEqual({
      deleteRange: [`sheetrows:${model.id}:${removed.id}:`, `sheetrows:${model.id}:${removed.id}:\uffff`],
    });
    expect(rowsWrites(idb).filter((record) => record.sheetId === removed.id)).toHaveLength(1);
    expect([...idb.data.keys()].some((key) => key.includes(removed.id))).toBe(false);
  });

  it("deleteDocument removes the shell, every rows record and the source record", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    store.deleteDocument(model.id);
    await flushWrites();
    expect([...idb.data.keys()].filter((key) => key.includes(model.id))).toEqual([]);
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
  });

  it("libraryStorageUsage counts the rows records", async () => {
    const { store, workbook } = await loadStore();
    const { rowsStoredBytes } = await import("./workbookStorage");
    const model = await workbook();
    store.storeSpreadsheet(model);
    await store.touchDocument(model.id);
    const withRows = store.libraryStorageUsage().totalBytes;
    const rowsBytes = model.sheets.reduce((total, sheet) => total + rowsStoredBytes(sheet.rows), 0);
    model.sheets = model.sheets.map((sheet) => ({ ...sheet, rows: [] }));
    await store.touchDocument(model.id);
    const withoutRows = store.libraryStorageUsage().totalBytes;
    expect(withRows - withoutRows).toBeGreaterThanOrEqual(rowsBytes - 64);
    expect(withRows).toBeGreaterThan(withoutRows);
  });

  it("a duplicate gets rows records of its own", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    clearWrites(idb);
    const copy = store.duplicateDocument(model.id)!;
    await flushWrites();
    expect(rowsWrites(idb).map((record) => record.docId)).toEqual([copy.id, copy.id, copy.id]);
  });
});

describe("view state (SPEC-117)", () => {
  it("storeSpreadsheetViewState writes the shell only, keeping lastModified — no rows record, nothing dirty", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    const lastModified = store.getStoredRecord(model.id)!.lastModified;
    clearWrites(idb);

    // A view-state change replaces the sheet objects' view fields; every
    // sheet keeps its rows array.
    const viewed: SpreadsheetModel = {
      ...model,
      activeSheetIndex: 1,
      sheets: model.sheets.map((sheet, index) =>
        index === 0
          ? {
              ...sheet,
              zoomPercent: 150,
              showGridLines: false,
              topLeftCell: { row: 40, col: 2 },
              selection: { active: { row: 41, col: 2 }, ranges: [{ top: 41, left: 2, bottom: 43, right: 3 }], activeRangeIndex: 0 },
            }
          : { ...sheet },
      ),
    };
    await store.storeSpreadsheetViewState(viewed);
    await flushWrites();

    expect(rowsWrites(idb)).toEqual([]);
    const shells = written(idb).filter((record) => record.id === model.id);
    expect(shells).toHaveLength(1);
    const shell = shells[0]!.model!;
    expect(shell.activeSheetIndex).toBe(1);
    expect(shell.sheets[0]).toMatchObject({ zoomPercent: 150, showGridLines: false, topLeftCell: { row: 40, col: 2 } });
    expect(shell.sheets[0]!.selection?.ranges).toEqual([{ top: 41, left: 2, bottom: 43, right: 3 }]);
    expect(shell.sheets.some((sheet) => "dirty" in sheet)).toBe(false);
    expect(store.getStoredRecord(model.id)!.lastModified).toBe(lastModified);
    expect(store.getSpreadsheet(model.id)).toBe(viewed);
    // The workbook's opaque package — what Download hands back — is untouched.
    expect(store.getSpreadsheet(model.id)!.opaque).toBe(model.opaque);
  });

  it("is a no-op for an id the store does not hold as a workbook", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    await store.storeSpreadsheetViewState(model);
    await flushWrites();
    expect(idb.writeBatch).not.toHaveBeenCalled();
    expect(idb.putRecord).not.toHaveBeenCalled();
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
  });
});

/** What IndexedDB holds after a workbook was persisted (read off a store
 * that wrote it). */
async function persisted() {
  const { store, idb, workbook } = await loadStore();
  const model = await workbook();
  store.storeSpreadsheet(model);
  await flushWrites();
  vi.resetModules();
  return { model, records: [...idb.data.values()] as Array<{ id: string }> };
}

describe("loadSpreadsheet rejoins the rows (SPEC-137's lazy residency)", () => {
  it("initLibrary reads the shell only; loadSpreadsheet rejoins each sheet's rows and seeds the latch (no rewrite on the next persist)", async () => {
    const { model, records } = await persisted();
    const { store, idb } = await loadStore(records);
    await store.initLibrary();
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
    expect(store.getSpreadsheetShell(model.id)!.sheets.map((sheet) => sheet.rows)).toEqual([[], [], []]);
    expect(store.listDocuments()).toHaveLength(1);

    const loaded = (await store.loadSpreadsheet(model.id))!;
    expect(loaded.sheets.map((sheet) => sheet.rows)).toEqual(model.sheets.map((sheet) => sheet.rows));
    store.renameDocument(model.id, "Renamed");
    await flushWrites();
    expect(rowsWrites(idb)).toEqual([]);
  });

  it("an orphaned chunk (a sheet the shell no longer has, still in the manifest) is deleted on the next persist", async () => {
    const { model, records } = await persisted();
    const shell = records.find((record) => record.id === model.id) as unknown as { sheetStore: { chunks: Record<string, number[]> } };
    shell.sheetStore.chunks.gone = [0];
    const orphan = { id: `sheetrows:${model.id}:gone:0000`, docId: model.id, sheetId: "gone", chunk: 0, rows: [] };
    const { store, idb } = await loadStore([...records, orphan]);
    await store.initLibrary();
    await store.loadSpreadsheet(model.id);
    await store.touchDocument(model.id);
    expect(idb.batches.at(-1)).toContainEqual({ deleteRange: [`sheetrows:${model.id}:gone:`, `sheetrows:${model.id}:gone:\uffff`] });
    expect(idb.data.has(orphan.id)).toBe(false);
  });

  it("refreshDocumentFromStorage adopts a newer stored record's shell; loadSpreadsheet rejoins its rows", async () => {
    const { model, records } = await persisted();
    const { store, idb } = await loadStore();
    for (const record of records) {
      idb.data.set(record.id, record);
    }
    const record = await store.refreshDocumentFromStorage(model.id);
    expect(store.isSpreadsheetModel(record!.model)).toBe(true);
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
    expect((await store.loadSpreadsheet(model.id))!.sheets[0]!.rows).toEqual(model.sheets[0]!.rows);
  });
});

describe("the reader-version upgrade on load", () => {
  /** A record an older reader stored: version 0, rows never extracted (no
   * rows records), the package retained, the user's title and star. */
  async function versionZeroRecord(): Promise<{ record: { id: string } & Record<string, unknown>; model: SpreadsheetModel }> {
    const { workbook } = await loadStore();
    const model = await workbook();
    vi.resetModules();
    const stored: SpreadsheetModel = {
      ...model,
      title: "Q3 numbers",
      readerVersion: 0,
      sheets: model.sheets.map((sheet, index) => ({ ...sheet, id: `kept-${index}`, rows: [] })),
    };
    delete stored.preview;
    return { model: stored, record: { id: stored.id, format: "xlsx", lastModified: 500, starred: true, model: stored } };
  }

  it("re-reads a version-0 record from opaque on load, keeping id, title, starred state and sheet ids, written back once", async () => {
    const { record, model } = await versionZeroRecord();
    const { store, idb } = await loadStore([record]);
    await store.initLibrary();
    const upgraded = (await store.loadSpreadsheet(model.id))!;

    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.id).toBe(model.id);
    expect(upgraded.title).toBe("Q3 numbers");
    expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(["kept-0", "kept-1", "kept-2"]);
    expect(upgraded.sheets[1]!.rows).toHaveLength(500);
    expect(store.getStoredRecord(model.id)!.starred).toBe(true);
    expect(store.getStoredRecord(model.id)!.lastModified).toBe(500);
    expect(store.listDocuments()[0]!.starred).toBe(true);

    await flushWrites();
    const shells = written(idb).filter((entry) => entry.id === model.id);
    expect(shells).toHaveLength(1);
    expect(shells[0]!.model!.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(rowsWrites(idb)).toHaveLength(3);
  });

  it("upgrades on load after a refresh too, and leaves a current record alone", async () => {
    const { record, model } = await versionZeroRecord();
    const { store, idb } = await loadStore();
    idb.data.set(record.id, record);
    await store.refreshDocumentFromStorage(model.id);
    const upgraded = (await store.loadSpreadsheet(model.id))!;
    expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(["kept-0", "kept-1", "kept-2"]);
    expect(await store.loadSpreadsheet(model.id)).toBe(upgraded);
  });

  it("keeps a record it cannot re-read exactly as it was", async () => {
    const { record, model } = await versionZeroRecord();
    model.opaque = [];
    const { store } = await loadStore([record]);
    await store.initLibrary();
    const loaded = (await store.loadSpreadsheet(model.id))!;
    expect(loaded.readerVersion).toBe(0);
    expect(loaded.title).toBe("Q3 numbers");
  });
});

describe("sheet commands write through (SPEC-121 AC-13)", () => {
  it("a rename of a sheet nobody references writes zero rows records; one that rewrites formulas writes exactly those sheets'", async () => {
    const { store, idb, workbook } = await loadStore();
    const { renameSheet } = await import("./spreadsheet/sheets");
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    clearWrites(idb);

    // "Hidden" is named by no formula: only the shell is written.
    renameSheet(model, 2, "Stash");
    await store.storeSpreadsheetEdit(model);
    expect(rowsWrites(idb)).toEqual([]);
    expect(written(idb).map((record) => record.id)).toEqual([model.id]);
    expect(written(idb)[0]!.model!.sheets[2]!.name).toBe("Stash");

    // A rename that rewrites a formula writes that sheet's rows and no other.
    const data = model.sheets[1]!;
    const summary = model.sheets[0]!;
    model.sheets = [{ ...summary, rows: [{ index: 9, cells: [{ col: 0, value: 1, formula: `${data.name}!A1` }] }] }, ...model.sheets.slice(1)];
    await store.storeSpreadsheetEdit(model);
    clearWrites(idb);
    renameSheet(model, 1, "Facts");
    await store.storeSpreadsheetEdit(model);
    expect(rowsWrites(idb).map((record) => record.sheetId)).toEqual([summary.id]);
  });
});
