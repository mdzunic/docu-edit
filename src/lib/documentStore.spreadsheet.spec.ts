// SPEC-116 in the store: a workbook is the fourth kind, and each sheet's rows
// live in their own `sheetrows:<docId>:<sheetId>` record — written only when
// that sheet's rows reference changed (the model's copy-on-write invariant),
// so a rename, a star or a library listing never structured-clones a cell;
// rejoined on every read path, deleted with the document (and with a sheet
// that no longer exists), counted by libraryStorageUsage. A record stored by
// an older reader is re-read from its retained package on load, keeping its
// id, title, starred state and sheet ids, and written back once.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { buildMiniBookXlsxBytes, buildSpecMacrosXlsmBytes } from "./spreadsheet/fixtures/build";
import type { Row, SpreadsheetModel } from "./spreadsheet/model";

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
  const { readSpreadsheetBytes } = await import("./spreadsheet/open");
  const workbook = async (): Promise<SpreadsheetModel> =>
    readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
  return { store, idb, workbook, readSpreadsheetBytes };
}

// The write-through is deferred to idle and yields a macrotask before each
// put (see documentStore.library.spec.ts): several real ticks catch up.
async function flushWrites() {
  for (let i = 0; i < 6; i++) {
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

function written(idb: ReturnType<typeof makeIdbMock>): WrittenRecord[] {
  return idb.putRecord.mock.calls.map((call) => call[0] as WrittenRecord);
}

function rowsWrites(idb: ReturnType<typeof makeIdbMock>): WrittenRecord[] {
  return written(idb).filter((record) => record.id.startsWith("sheetrows:"));
}

function deleted(idb: ReturnType<typeof makeIdbMock>): string[] {
  return idb.deleteRecord.mock.calls.map((call) => String(call[0]));
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

describe("the per-sheet row split", () => {
  it("writes each sheet's rows to its own record and the shell with every sheet's rows emptied", async () => {
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
      model.sheets.map((sheet) => `sheetrows:${model.id}:${sheet.id}`).toSorted(),
    );
    for (const sheet of model.sheets) {
      const record = rows.find((entry) => entry.sheetId === sheet.id)!;
      expect(record.docId).toBe(model.id);
      expect(record.rows).toBe(sheet.rows);
    }
    // The in-memory model keeps its rows.
    expect(store.getSpreadsheet(model.id)!.sheets[1]!.rows).toHaveLength(500);
  });

  it("a rename, a star or a library listing writes no rows record", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    idb.putRecord.mockClear();

    store.renameDocument(model.id, "Budget");
    store.setStarred(model.id, true);
    store.listDocuments();
    store.libraryStorageUsage();
    await flushWrites();

    expect(rowsWrites(idb)).toEqual([]);
    expect(written(idb).map((record) => record.id)).toEqual([model.id, model.id]);
    expect(written(idb).at(-1)!.model!.title).toBe("Budget");
  });

  it("editing one sheet's rows array writes exactly that sheet's record, and refreshes the preview", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    idb.putRecord.mockClear();

    // Copy-on-write: a new A1 cell, a new first row, a new rows array; every
    // other row keeps its identity.
    const summary = model.sheets[0]!;
    const [first, ...rest] = summary.rows;
    summary.rows = [{ ...first!, cells: [{ col: 0, value: "Edited" }, ...first!.cells.slice(1)] }, ...rest];
    await store.touchDocument(model.id);

    const rows = rowsWrites(idb);
    expect(rows.map((record) => record.id)).toEqual([`sheetrows:${model.id}:${summary.id}`]);
    expect(rows[0]!.rows).toBe(summary.rows);
    const shell = written(idb).find((record) => record.id === model.id)!;
    expect(shell.model!.preview?.cells[0]?.[0]).toBe("Edited");
  });

  it("deletes a removed sheet's record on the next persist", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    const removed = model.sheets[2]!;
    model.sheets = model.sheets.slice(0, 2);
    await store.touchDocument(model.id);
    expect(deleted(idb)).toContain(`sheetrows:${model.id}:${removed.id}`);
    expect(rowsWrites(idb).filter((record) => record.sheetId === removed.id)).toHaveLength(1);
  });

  it("deleteDocument removes the shell and every rows record", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    store.deleteDocument(model.id);
    const gone = deleted(idb);
    expect(gone).toContain(model.id);
    for (const sheet of model.sheets) {
      expect(gone).toContain(`sheetrows:${model.id}:${sheet.id}`);
    }
    expect(store.getSpreadsheet(model.id)).toBeUndefined();
  });

  it("libraryStorageUsage counts the rows records", async () => {
    const { store, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    const withRows = store.libraryStorageUsage().totalBytes;
    const rowsBytes = model.sheets.reduce((total, sheet) => total + JSON.stringify(sheet.rows).length, 0);
    for (const sheet of model.sheets) {
      sheet.rows = [];
    }
    const withoutRows = store.libraryStorageUsage().totalBytes;
    expect(withRows - withoutRows).toBeGreaterThanOrEqual(rowsBytes - 3 * "[]".length);
    expect(withRows).toBeGreaterThan(withoutRows);
  });

  it("a duplicate gets rows records of its own", async () => {
    const { store, idb, workbook } = await loadStore();
    const model = await workbook();
    store.storeSpreadsheet(model);
    await flushWrites();
    idb.putRecord.mockClear();
    const copy = store.duplicateDocument(model.id)!;
    await flushWrites();
    expect(rowsWrites(idb).map((record) => record.docId)).toEqual([copy.id, copy.id, copy.id]);
  });
});

/** What IndexedDB holds after a workbook was persisted: the shell plus a
 * rows record per sheet. */
function persisted(model: SpreadsheetModel, extra: Record<string, unknown> = {}) {
  const shell = {
    id: model.id,
    format: model.format,
    lastModified: 1000,
    ...extra,
    model: { ...model, media: [], sheets: model.sheets.map((sheet) => ({ ...sheet, rows: [] })) },
  };
  const rows = model.sheets.map((sheet) => ({
    id: `sheetrows:${model.id}:${sheet.id}`,
    docId: model.id,
    sheetId: sheet.id,
    rows: sheet.rows,
  }));
  return { shell, rows };
}

describe("every read path rejoins the rows", () => {
  it("initLibrary rejoins each sheet's rows and seeds the latch (no rewrite on the next persist)", async () => {
    const { workbook } = await loadStore();
    const model = await workbook();
    vi.resetModules();
    const { shell, rows } = persisted(model);
    const { store, idb } = await loadStore([shell, ...rows]);
    await store.initLibrary();

    const loaded = store.getSpreadsheet(model.id)!;
    expect(loaded.sheets.map((sheet) => sheet.rows)).toEqual(model.sheets.map((sheet) => sheet.rows));
    expect(store.listDocuments()).toHaveLength(1);
    store.renameDocument(model.id, "Renamed");
    await flushWrites();
    expect(rowsWrites(idb)).toEqual([]);
  });

  it("an orphaned rows record (a sheet the shell no longer has) is deleted on the next persist", async () => {
    const { workbook } = await loadStore();
    const model = await workbook();
    vi.resetModules();
    const { shell, rows } = persisted(model);
    const orphan = { id: `sheetrows:${model.id}:gone`, docId: model.id, sheetId: "gone", rows: [] };
    const { store, idb } = await loadStore([shell, ...rows, orphan]);
    await store.initLibrary();
    await store.touchDocument(model.id);
    expect(deleted(idb)).toEqual([orphan.id]);
  });

  it("refreshDocumentFromStorage rejoins the rows of a newer stored record", async () => {
    const { workbook } = await loadStore();
    const model = await workbook();
    vi.resetModules();
    const { shell, rows } = persisted(model);
    const { store, idb } = await loadStore();
    idb.getRecord.mockImplementation(async (key: string) => (key === model.id ? shell : rows.find((record) => record.id === key)));
    const record = await store.refreshDocumentFromStorage(model.id);
    expect(store.isSpreadsheetModel(record!.model)).toBe(true);
    expect(store.getSpreadsheet(model.id)!.sheets[0]!.rows).toEqual(model.sheets[0]!.rows);
  });
});

describe("the reader-version upgrade on load", () => {
  /** A record an older reader stored: version 0, rows never extracted (no
   * rows records), the package retained, the user's title and star. */
  async function versionZeroRecord(): Promise<{ record: Record<string, unknown>; model: SpreadsheetModel }> {
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

  it("re-reads a version-0 record from opaque on initLibrary, keeping id, title, starred state and sheet ids, written back once", async () => {
    const { record, model } = await versionZeroRecord();
    const { store, idb } = await loadStore([record]);
    await store.initLibrary();

    const upgraded = store.getSpreadsheet(model.id)!;
    expect(upgraded.readerVersion).toBe(1);
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
    expect(shells[0]!.model!.readerVersion).toBe(1);
    expect(rowsWrites(idb)).toHaveLength(3);
  });

  it("upgrades on the refresh path too, and leaves a current record alone", async () => {
    const { record, model } = await versionZeroRecord();
    const { store, idb } = await loadStore();
    idb.getRecord.mockImplementation(async (key: string) => (key === model.id ? record : undefined));
    await store.refreshDocumentFromStorage(model.id);
    const upgraded = store.getSpreadsheet(model.id)!;
    expect(upgraded.readerVersion).toBe(1);
    expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(["kept-0", "kept-1", "kept-2"]);
    expect(await store.upgradeStoredSpreadsheet(model.id)).toBe(upgraded);
  });

  it("keeps a record it cannot re-read exactly as it was", async () => {
    const { record, model } = await versionZeroRecord();
    model.opaque = [];
    const { store } = await loadStore([record]);
    await store.initLibrary();
    expect(store.getSpreadsheet(model.id)!.readerVersion).toBe(0);
    expect(store.getSpreadsheet(model.id)!.title).toBe("Q3 numbers");
  });
});
