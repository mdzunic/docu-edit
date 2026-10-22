// SPEC-121 AC-1: lib/newDocument.ts is the one place the landing page's
// create actions live. Each kind stores its blank model through its own
// store accessor, carrying its format from birth, and returns the editor path
// documentKinds routes that kind to. Fresh modules per test (the store is
// module state) with the idb seam mocked, as in
// documentStore.spreadsheet.spec.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { editorPathFor } from "./documentKinds";
import { DEFAULT_TITLE } from "./documentModel";
import { DEFAULT_PRESENTATION_TITLE } from "./presentation/model";
import { DEFAULT_SPREADSHEET_TITLE } from "./spreadsheet/model";

function makeIdbMock() {
  return {
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    getRecord: vi.fn().mockResolvedValue(undefined),
  };
}

async function loadModules() {
  const idb = makeIdbMock();
  vi.doMock("./idb", () => idb);
  const store = await import("./documentStore");
  const { savePreference } = await import("./preferences");
  const { BLANK_DOCUMENT_KINDS, createBlankDocument } = await import("./newDocument");
  return { idb, store, savePreference, BLANK_DOCUMENT_KINDS, createBlankDocument };
}

// The write-through is deferred to idle and yields a macrotask before each
// put (see documentStore.library.spec.ts): several real ticks catch up.
async function flushWrites() {
  for (let i = 0; i < 6; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

/** The id an editor path ends in: "/sheets/<id>" → "<id>". */
function idOf(path: string): string {
  return path.slice(path.lastIndexOf("/") + 1);
}

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
});

afterEach(() => {
  vi.doUnmock("./idb");
  vi.restoreAllMocks();
});

describe("createBlankDocument (SPEC-121 AC-1)", () => {
  it('"text" stores a blank text document (format "new", the new-document preferences applied) and returns /doc/:id', async () => {
    const { store, savePreference, createBlankDocument } = await loadModules();
    savePreference("showRuler", false);

    const path = createBlankDocument("text");

    expect(path).toMatch(/^\/doc\/[^/]+$/);
    const doc = store.getDocument(idOf(path));
    expect(doc?.title).toBe(DEFAULT_TITLE);
    // Through getOrCreateDocument, exactly as "New document" always was: the
    // SPEC-066 "new documents" preferences seed the model.
    expect(doc?.showRuler).toBe(false);
    expect(store.getStoredRecord(idOf(path))?.format).toBe("new");
  });

  it('"presentation" stores a blank deck, "pptx" from birth, and returns /slides/:id', async () => {
    const { store, createBlankDocument } = await loadModules();

    const path = createBlankDocument("presentation");

    expect(path).toMatch(/^\/slides\/[^/]+$/);
    const deck = store.getPresentation(idOf(path));
    expect(deck?.title).toBe(DEFAULT_PRESENTATION_TITLE);
    expect(deck?.format).toBe("pptx");
    expect(deck?.slides).toHaveLength(1);
    expect(store.getStoredRecord(idOf(path))?.format).toBe("pptx");
  });

  it('"spreadsheet" stores a blank workbook, "xlsx" from birth with the one sheet "Sheet1", and returns /sheets/:id', async () => {
    const { store, createBlankDocument } = await loadModules();

    const path = createBlankDocument("spreadsheet");

    expect(path).toMatch(/^\/sheets\/[^/]+$/);
    const workbook = store.getSpreadsheet(idOf(path));
    expect(DEFAULT_SPREADSHEET_TITLE).toBe("Untitled spreadsheet");
    expect(workbook?.title).toBe(DEFAULT_SPREADSHEET_TITLE);
    expect(workbook?.format).toBe("xlsx");
    expect(workbook?.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1"]);
    expect(store.getStoredRecord(idOf(path))?.format).toBe("xlsx");
    // Stored through storeSpreadsheet, which gives the library card its
    // miniature from the active sheet.
    expect(workbook?.preview?.sheetName).toBe("Sheet1");
  });

  it("lists each new document under its own kind and returns the route documentKinds gives that kind", async () => {
    const { store, BLANK_DOCUMENT_KINDS, createBlankDocument } = await loadModules();
    expect(BLANK_DOCUMENT_KINDS).toEqual(["text", "presentation", "spreadsheet"]);

    for (const kind of BLANK_DOCUMENT_KINDS) {
      const path = createBlankDocument(kind);
      const id = idOf(path);
      expect(store.documentKindOf(id)).toBe(kind);
      expect(path).toBe(editorPathFor(kind, id));
    }
    expect(store.listDocuments().map((entry) => entry.kind).toSorted()).toEqual([
      "presentation",
      "spreadsheet",
      "text",
    ]);
  });

  it("mints a fresh id on every call: a second New spreadsheet is a second workbook", async () => {
    const { store, createBlankDocument } = await loadModules();

    const first = createBlankDocument("spreadsheet");
    const second = createBlankDocument("spreadsheet");

    expect(second).not.toBe(first);
    expect(store.listDocuments().map((entry) => entry.id).toSorted()).toEqual([idOf(first), idOf(second)].toSorted());
  });

  it("writes the new workbook through to storage like any other record", async () => {
    const { idb, createBlankDocument } = await loadModules();

    const id = idOf(createBlankDocument("spreadsheet"));
    await flushWrites();

    const written = idb.putRecord.mock.calls.map((call) => (call[0] as { id: string }).id);
    expect(written).toContain(id);
  });
});
