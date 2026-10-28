// SPEC-138 — the workbook's saving: useAutosave's state machine reused for
// workbooks (drafts, commits, recovery's Keep/Discard with the workbook's
// sheetView) behind SPEC-010's edit lock — injected here as a LockManagerLike,
// so a second tab's "busy" and the queued grant are steps a test takes. The
// idb seam is an in-memory double of the documents store, so what gets
// written — and what never does — is observable.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h, ref, shallowRef } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import type { LockManagerLike } from "../../lib/core/editLock";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetViewState } from "../../lib/spreadsheet/selection";

const inRange = (key: string, lower: string, upper: string) => key >= lower && key <= upper;

function memoryIdb() {
  const documents = new Map<string, { id: string }>();
  const puts: string[] = [];
  const write = (record: { id: string }) => {
    puts.push(record.id);
    documents.set(record.id, structuredClone(record));
  };
  return {
    documents,
    puts,
    loadAllRecords: vi.fn(async () => [...documents.values()].map((r) => structuredClone(r))),
    getRecord: vi.fn(async (id: string) => structuredClone(documents.get(id))),
    putRecord: vi.fn(async (r: { id: string }) => write(r)),
    deleteRecord: vi.fn(async (id: string) => void documents.delete(id)),
    deleteRecordRange: vi.fn(async (lower: string, upper: string) => {
      for (const key of documents.keys()) {
        if (inRange(key, lower, upper)) {
          documents.delete(key);
        }
      }
    }),
    writeBatch: vi.fn(async (ops: Array<{ put?: { id: string }; delete?: string; deleteRange?: [string, string] }>) => {
      for (const op of ops) {
        if (op.put) {
          write(op.put);
        } else if (op.delete !== undefined) {
          documents.delete(op.delete);
        } else if (op.deleteRange) {
          for (const key of documents.keys()) {
            if (inRange(key, op.deleteRange[0], op.deleteRange[1])) {
              documents.delete(key);
            }
          }
        }
      }
    }),
    getVersionsRecord: vi.fn(async () => undefined),
    putVersionsRecord: vi.fn(async () => undefined),
    deleteVersionsRecord: vi.fn(async () => undefined),
  };
}

let idb: ReturnType<typeof memoryIdb>;
const wrappers: VueWrapper[] = [];

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  idb = memoryIdb();
  vi.doMock("../../lib/idb", () => idb);
});

afterEach(() => {
  for (const wrapper of wrappers.splice(0)) {
    wrapper.unmount();
  }
  vi.doUnmock("../../lib/idb");
});

/** Real macrotask turns: the store yields one before every write. */
async function settle(turns = 30): Promise<void> {
  for (let i = 0; i < turns; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

/** A lock manager a test drives: "busy" while another tab holds the lock,
 * then `grant()` hands the queued request the lock, as the browser does
 * when that tab closes. */
function controlledLocks(busy: boolean) {
  let queued: ((lock: unknown) => unknown) | null = null;
  let released = 0;
  const locks: LockManagerLike = {
    request(name, options, callback) {
      if (options.ifAvailable) {
        const held = callback(busy ? null : { name });
        void Promise.resolve(held).then(() => (released += busy ? 0 : 1));
        return Promise.resolve(held);
      }
      queued = callback;
      return new Promise(() => {});
    },
  };
  return {
    locks,
    grant: () => {
      const held = queued?.({ name: "queued" });
      void Promise.resolve(held).then(() => (released += 1));
    },
    released: () => released,
  };
}

function rows(count: number, tag: string): Row[] {
  return Array.from({ length: count }, (_, index) => ({ index, cells: [{ col: 0, value: `${tag}${index}` }] }));
}

async function workbook(): Promise<SpreadsheetModel> {
  const { createBlankWorkbook } = await import("../../lib/spreadsheet/sheets");
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  model.sheets = [
    { ...base, id: "one", name: "One", rows: rows(4, "a") },
    { ...base, id: "two", name: "Two", rows: rows(4, "b") },
  ];
  return model;
}

const COMMITTED_VIEW: SheetViewState = { sheetId: "one", active: { row: 1, col: 0 }, topRow: 0, leftCol: 0 };
const DRAFT_VIEW: SheetViewState = { sheetId: "two", active: { row: 3, col: 0 }, topRow: 2, leftCol: 0 };

async function harness(options: { locks?: LockManagerLike; draft?: boolean; beforeMount?: () => void } = {}) {
  const store = await import("../../lib/documentStore");
  const { useWorkbookSaving } = await import("./useWorkbookSaving");
  const committed = await workbook();
  store.storeSpreadsheet(committed);
  await store.commitDocument(committed.id, { caret: null, scroll: 0, sheetView: COMMITTED_VIEW });
  if (options.draft) {
    // What a crashed session leaves: an edit on "Two", drafted after the commit.
    await new Promise((resolve) => setTimeout(resolve, 5));
    const edited: SpreadsheetModel = { ...committed, sheets: committed.sheets.map((s, i) => (i === 1 ? { ...s, rows: rows(4, "drafted") } : s)) };
    store.replaceSpreadsheetModel(committed.id, edited);
    await store.saveDraft(committed.id, { caret: null, scroll: 0, sheetView: DRAFT_VIEW });
    store.replaceSpreadsheetModel(committed.id, committed);
  }
  // The import's own idle write-through lands before the log starts.
  await settle();
  idb.puts.length = 0;
  options.beforeMount?.();
  const model = shallowRef<SpreadsheetModel | null>(committed);
  const show = vi.fn((next: SpreadsheetModel) => {
    model.value = next;
  });
  const applySheetView = vi.fn(() => true);
  const resetUndo = vi.fn();
  const remount = vi.fn();
  const download = vi.fn();
  let saving: ReturnType<typeof useWorkbookSaving> | null = null;
  const wrapper = mount(
    defineComponent({
      setup() {
        saving = useWorkbookSaving({
          model,
          workbookId: ref(committed.id),
          getSheetView: () => DRAFT_VIEW,
          applySheetView,
          resetUndo,
          download,
          show,
          remount,
          locks: options.locks,
        });
        return () => h("div");
      },
    }),
  );
  wrappers.push(wrapper);
  await settle();
  return { saving: saving!, store, committed, model, show, applySheetView, resetUndo, remount, wrapper };
}

describe("the edit lock (AC-32..AC-35)", () => {
  it("a second tab opens read-only: no session, no badge, no edit noted — until the queued lock grants", async () => {
    const lock = controlledLocks(true);
    const { saving, store, committed, model, resetUndo } = await harness({ locks: lock.locks });
    expect(saving.lockState.value).toBe("busy");
    expect(saving.conflictMessage.value).toBe("This workbook is open in another tab");
    expect(saving.editable.value).toBe(false);
    expect(saving.showSaveState.value).toBe(false);
    expect(idb.getRecord.mock.calls.map(([key]) => key)).not.toContain(`draft:${committed.id}`);
    // An edit attempt notes nothing: the store keeps its model, nothing is written.
    model.value = { ...committed, title: "Typed in the busy tab" };
    saving.noteEdit();
    await settle();
    expect(store.getStoredRecord(committed.id)!.model.title).not.toBe("Typed in the busy tab");
    expect(idb.puts).toEqual([]);

    // The other tab closes: the queued lock grants; the session begins over
    // the stored commit, the undo stack resets and editing enables.
    model.value = committed;
    lock.grant();
    await settle();
    expect(saving.lockState.value).toBe("acquired");
    expect(saving.conflictMessage.value).toBe("");
    expect(saving.editable.value).toBe(true);
    expect(saving.showSaveState.value).toBe(true);
    expect(resetUndo).toHaveBeenCalledTimes(1);
    expect(idb.getRecord.mock.calls.map(([key]) => key)).toContain(`draft:${committed.id}`);
  });

  it("a first tab acquires at once — nothing stacked, so no undo reset — and releases the lock on unmount", async () => {
    const lock = controlledLocks(false);
    const { saving, resetUndo, wrapper } = await harness({ locks: lock.locks });
    expect(saving.lockState.value).toBe("acquired");
    expect(resetUndo).not.toHaveBeenCalled();
    expect(lock.released()).toBe(0);
    wrapper.unmount();
    wrappers.length = 0;
    await settle();
    expect(lock.released()).toBe(1);
  });
});

describe("edits and commit points (AC-9, AC-10)", () => {
  it("noteEdit hands the live model to the store and arms the draft cadence; a commit point commits and tells its listeners", async () => {
    const { saving, store, committed, model } = await harness();
    const heard: string[] = [];
    saving.listen({ edit: () => heard.push("edit"), commitPoint: () => heard.push("commit") });
    const edited: SpreadsheetModel = { ...committed, sheets: committed.sheets.map((s, i) => (i === 0 ? { ...s, rows: rows(4, "edited") } : s)) };
    model.value = edited;
    saving.noteEdit();
    expect(store.getStoredRecord(committed.id)!.model).toBe(edited);
    expect(saving.saveState.value).toBe("saving");
    expect(heard).toEqual(["edit"]);
    // Nothing was written through.
    expect(idb.puts).toEqual([]);

    saving.commitPoint();
    await settle();
    expect(saving.saveState.value).toBe("saved");
    expect(heard).toEqual(["edit", "commit"]);
    expect(idb.puts).toContain(`sheetrows:${committed.id}:one:0000`);
    expect(idb.puts).toContain(committed.id);
    expect((idb.documents.get(committed.id) as unknown as { sheetView: SheetViewState }).sheetView).toEqual(DRAFT_VIEW);
  });

  it("flushForClose commits before the view goes and reports a clean close", async () => {
    const { saving, committed, model } = await harness();
    model.value = { ...committed, title: "Closing" };
    saving.noteEdit();
    expect(await saving.flushForClose()).toBe(true);
    expect((idb.documents.get(committed.id) as unknown as { model: { title: string } }).model.title).toBe("Closing");
  });
});

describe("recovery and the sheetView (AC-26..AC-29)", () => {
  it("reopens over a newer draft: the recovered workbook shown on its drafted view; Discard brings back the committed one and its view, and resets undo", async () => {
    const { saving, committed, model, show, resetUndo } = await harness({ draft: true });
    expect(saving.recovery.value).not.toBeNull();
    expect(saving.recoveryLabel.value).toMatch(/^Recovered unsaved changes from/);
    expect(saving.saveStateText.value).toBe("Not saved yet");
    const recovered = model.value!;
    expect(recovered).not.toBe(committed);
    expect(recovered.id).toBe(committed.id);
    expect(recovered.sheets[1]!.rows[0]!.cells[0]!.value).toBe("drafted0");
    // The drafted view: its sheet active, its active cell the selection, its top-left cell.
    expect(recovered.activeSheetIndex).toBe(1);
    expect(recovered.sheets[1]!.selection?.active).toEqual({ row: 3, col: 0 });
    expect(recovered.sheets[1]!.topLeftCell).toEqual({ row: 2, col: 0 });
    expect(show).toHaveBeenCalledWith(recovered);

    resetUndo.mockClear();
    saving.discardRecovered();
    expect(saving.recovery.value).toBeNull();
    expect(model.value).toBe(committed);
    expect(committed.activeSheetIndex).toBe(0);
    expect(committed.sheets[0]!.selection?.active).toEqual({ row: 1, col: 0 });
    expect(resetUndo).toHaveBeenCalledTimes(1);
    await settle();
    expect(idb.documents.has(`draft:${committed.id}`)).toBe(false);
  });

  it("Keep commits the recovered workbook and hides the bar", async () => {
    const { saving, committed } = await harness({ draft: true });
    await saving.keepRecovered();
    await settle();
    expect(saving.recovery.value).toBeNull();
    expect(saving.saveStateText.value).toBe("All changes saved");
    const stored = idb.documents.get(`sheetrows:${committed.id}:two:0000`) as unknown as { rows: Row[] };
    expect(stored.rows[0]!.cells[0]!.value).toBe("drafted0");
    expect(idb.documents.has(`draft:${committed.id}`)).toBe(false);
  });

  it("a plain reopen applies the record's sheetView to the model on show, re-mounting the grid", async () => {
    const { applySheetView, remount } = await harness();
    expect(applySheetView).toHaveBeenCalledWith(COMMITTED_VIEW);
    expect(remount).toHaveBeenCalledTimes(1);
  });

  it("withSheetView ignores a view naming a sheet that no longer exists", async () => {
    const { withSheetView } = await import("./useSheetViewState");
    const model = await workbook();
    expect(withSheetView(model, { sheetId: "gone", active: { row: 0, col: 0 }, topRow: 0, leftCol: 0 })).toBe(false);
    expect(model.activeSheetIndex).toBe(0);
    expect(withSheetView(model, DRAFT_VIEW)).toBe(true);
    expect(model.activeSheetIndex).toBe(1);
    expect(model.sheets[1]!.rows).toEqual(rows(4, "b"));
  });
});

describe("the session's begin (review: nothing commits or edits before it)", () => {
  it("while the crash draft is still being read, a commit point and a close commit nothing and edits wait — then the bar offers it", async () => {
    const opener: { open?: () => void } = {};
    const gate = new Promise<void>((resolve) => {
      opener.open = resolve;
    });
    const { saving, committed } = await harness({
      draft: true,
      beforeMount: () => {
        const read = idb.getRecord.getMockImplementation()!;
        idb.getRecord.mockImplementation(async (key: string) => {
          if (key.startsWith("draft:")) {
            await gate;
          }
          return read(key);
        });
      },
    });
    expect(saving.lockState.value).toBe("acquired");
    expect(saving.editable.value).toBe(false);
    expect(saving.owns.value).toBe(false);
    saving.commitPoint();
    expect(await saving.flushForClose()).toBe(true);
    await settle();
    expect(idb.puts).toEqual([]);
    expect(idb.documents.has(`draft:${committed.id}`)).toBe(true);

    opener.open!();
    await settle();
    expect(saving.recovery.value).not.toBeNull();
    expect(saving.editable.value).toBe(true);
  });

  it("Discard while a version is previewed ends the preview before the committed workbook is shown", async () => {
    const { saving, committed, model } = await harness({ draft: true });
    const order: string[] = [];
    saving.listen({ swap: () => order.push(`swap over ${model.value === committed ? "committed" : "other"}`) });
    saving.previewing.value = true;
    saving.discardRecovered();
    expect(order).toEqual(["swap over other"]);
    expect(model.value).toBe(committed);
  });

  it("a restore's commit settles an open recovery bar, committing what is on show", async () => {
    const { saving, store, committed } = await harness({ draft: true });
    const restored: SpreadsheetModel = { ...committed, sheets: committed.sheets.map((s, i) => (i === 0 ? { ...s, rows: rows(4, "restored") } : s)) };
    store.replaceSpreadsheetModel(committed.id, restored);
    await saving.commitNow();
    await settle();
    expect(saving.recovery.value).toBeNull();
    const stored = idb.documents.get(`sheetrows:${committed.id}:one:0000`) as unknown as { rows: Row[] };
    expect(stored.rows[0]!.cells[0]!.value).toBe("restored0");
    expect(idb.documents.has(`draft:${committed.id}`)).toBe(false);
  });
});
