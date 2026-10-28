// SPEC-138 — the workbook's version history controller: the snapshot
// triggers it hears from the saving session (an edit arms the five-minute
// cadence, a commit point snapshots a changed workbook, Keep version history
// gates only those two), preview in the grid without touching the live
// workbook or the store, restore as a replace — pre-restore, the swap, an
// empty undo stack, the commit, post-restore, the toast whose "Undo restore"
// restores the pre-restore version the same way — the panel's labels, its
// paused notices and its read-only state in a tab that does not own the
// workbook. The idb seam is an in-memory double of both object stores.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, defineComponent, h, ref, shallowRef } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import type { Row, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { WorkbookSessionListener } from "./useWorkbookSaving";

function memoryIdb() {
  const documents = new Map<string, { id: string }>();
  const versions = new Map<string, { id: string }>();
  return {
    documents,
    versions,
    loadAllRecords: vi.fn(async () => []),
    getRecord: vi.fn(async (id: string) => structuredClone(documents.get(id))),
    putRecord: vi.fn(async (r: { id: string }) => void documents.set(r.id, structuredClone(r))),
    deleteRecord: vi.fn(async (id: string) => void documents.delete(id)),
    deleteRecordRange: vi.fn(async () => undefined),
    writeBatch: vi.fn(async (ops: Array<{ put?: { id: string } }>) => {
      for (const op of ops) {
        if (op.put) {
          documents.set(op.put.id, structuredClone(op.put));
        }
      }
    }),
    getVersionsRecord: vi.fn(async (id: string) => structuredClone(versions.get(id))),
    putVersionsRecord: vi.fn(async (r: { id: string }) => void versions.set(r.id, structuredClone(r))),
    deleteVersionsRecord: vi.fn(async (id: string) => void versions.delete(id)),
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
  vi.useRealTimers();
});

async function settle(turns = 30): Promise<void> {
  for (let i = 0; i < turns; i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

function rows(count: number, tag: string): Row[] {
  return Array.from({ length: count }, (_, index) => ({ index, cells: [{ col: 0, value: `${tag}${index}` }] }));
}

async function workbook(tag = "a"): Promise<SpreadsheetModel> {
  const { createBlankWorkbook } = await import("../../lib/spreadsheet/sheets");
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  model.sheets = [
    { ...base, id: "one", name: "One", rows: rows(3, tag) },
    { ...base, id: "two", name: "Two", rows: rows(2, "b") },
  ];
  model.activeSheetIndex = 1;
  return model;
}

/** An edit as the pipeline makes one: a new rows array on a new sheet. */
function edit(model: SpreadsheetModel, tag: string): void {
  model.sheets = model.sheets.map((sheet, index) => (index === 0 ? { ...sheet, rows: rows(3, tag), dirty: true } : sheet));
}

async function harness(options: { owns?: boolean } = {}) {
  const store = await import("../../lib/documentStore");
  const { useWorkbookHistory } = await import("./useWorkbookHistory");
  const { useSheetSidePanel } = await import("./useSheetSidePanel");
  const live = await workbook();
  store.storeSpreadsheet(live);
  await settle();
  const model = shallowRef<SpreadsheetModel | null>(live);
  const listeners: WorkbookSessionListener[] = [];
  const owns = ref(options.owns ?? true);
  const previewing = ref(false);
  const commitNow = vi.fn(async () => store.commitDocument(live.id, { caret: null, scroll: 0 }));
  const shown: SpreadsheetModel[] = [];
  const show = vi.fn((next: SpreadsheetModel) => {
    shown.push(next);
    model.value = next;
  });
  const syncView = vi.fn();
  const resetUndo = vi.fn();
  let history: ReturnType<typeof useWorkbookHistory> | null = null;
  let sidePanel: ReturnType<typeof useSheetSidePanel> | null = null;
  const wrapper = mount(
    defineComponent({
      setup() {
        sidePanel = useSheetSidePanel();
        history = useWorkbookHistory({
          saving: {
            listen: (listener: WorkbookSessionListener) => void listeners.push(listener),
            owns: computed(() => owns.value),
            previewing,
            commitNow,
          },
          model,
          show,
          syncView,
          resetUndo,
          sidePanel,
        });
        return () => h("div");
      },
    }),
  );
  wrappers.push(wrapper);
  const emit = (event: keyof WorkbookSessionListener, id = live.id) => listeners.forEach((listener) => listener[event]?.(id));
  emit("begin");
  await settle();
  return { history: history!, sidePanel: sidePanel!, store, live, model, owns, previewing, commitNow, show, shown, syncView, resetUndo, emit };
}

function versionTriggers(): (string | undefined)[] {
  const index = [...idb.versions.values()].find((r) => (r as { store?: string }).store === "index") as { versions: { trigger?: string }[] } | undefined;
  return index?.versions.map((v) => v.trigger) ?? [];
}

describe("snapshot triggers (AC-42, AC-43)", () => {
  it("a commit point snapshots a changed workbook only; an edit arms the five-minute cadence", async () => {
    const { history, live, emit } = await harness();
    emit("commitPoint");
    await settle();
    expect(versionTriggers()).toEqual([]);
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    edit(live, "edited");
    emit("edit");
    expect(vi.getTimerCount()).toBe(1);
    emit("commitPoint");
    await vi.waitFor(() => expect(versionTriggers()).toEqual(["commit"]));
    // More editing: the cadence already armed is not re-armed, and fires
    // five minutes after the first edit.
    edit(live, "again");
    emit("edit");
    expect(vi.getTimerCount()).toBe(1);
    await vi.advanceTimersByTimeAsync(5 * 60 * 1000);
    expect(vi.getTimerCount()).toBe(0);
    vi.useRealTimers();
    await vi.waitFor(() => expect(versionTriggers()).toEqual(["commit", "cadence"]));
    await vi.waitFor(() => expect(history.versions.value.map((v) => v.trigger)).toEqual(["cadence", "commit"]));
  });

  it("Keep version history off stops cadence and commit snapshots, never a restore's pair", async () => {
    const { history, live, emit } = await harness();
    const { usePreferences } = await import("../../composables/usePreferences");
    const { recordWorkbookVersion } = await import("../../lib/spreadsheetVersions");
    const seeded = await recordWorkbookVersion({ docId: live.id, model: await workbook("seeded"), author: "You", trigger: "commit", cachedPrev: null });
    usePreferences().setPreference("keepVersionHistory", false);
    edit(live, "edited");
    emit("edit");
    emit("commitPoint");
    await settle();
    expect(versionTriggers()).toEqual(["commit"]);
    await history.restore(seeded.newEntry!.id);
    expect(versionTriggers()).toEqual(["commit", "pre-restore", "post-restore"]);
  });
});

describe("restore is a replace (AC-54, AC-55)", () => {
  it("pre-restore → reconstruct → replace → show → empty undo → commit → post-restore → the toast; Undo restore restores the pre-restore version the same way", async () => {
    const { history, store, live, model, commitNow, resetUndo } = await harness();
    const { recordWorkbookVersion } = await import("../../lib/spreadsheetVersions");
    const { useToasts } = await import("../../composables/useToasts");
    const old = { ...(await workbook("old")), id: live.id, title: "Old title" };
    const target = await recordWorkbookVersion({ docId: live.id, model: old, author: "You", trigger: "commit", cachedPrev: null });
    edit(live, "live");

    await history.restore(target.newEntry!.id);
    const restored = model.value!;
    expect(restored).not.toBe(live);
    expect(restored.id).toBe(live.id);
    expect(restored.title).toBe("Old title");
    expect(restored.sheets[0]!.rows[0]!.cells[0]!.value).toBe("old0");
    expect(store.getSpreadsheet(live.id)).toBe(restored);
    expect(resetUndo).toHaveBeenCalledTimes(1);
    expect(commitNow).toHaveBeenCalledTimes(1);
    expect(versionTriggers().slice(-2)).toEqual(["pre-restore", "post-restore"]);
    const toast = useToasts().toasts.value.at(-1)!;
    expect(toast.message).toBe("Version restored");
    expect(toast.action?.label).toBe("Undo restore");

    // "Undo restore": the pre-restore version back, the same way — two more snapshots.
    toast.action!.run();
    await vi.waitFor(() => expect(model.value!.sheets[0]!.rows[0]!.cells[0]!.value).toBe("live0"));
    await vi.waitFor(() => expect(versionTriggers()).toHaveLength(5));
    expect(resetUndo).toHaveBeenCalledTimes(2);
  });

  it("a version whose blob is missing can't be opened: the panel says so and nothing changes", async () => {
    const { history, live, model, resetUndo } = await harness();
    const { recordWorkbookVersion } = await import("../../lib/spreadsheetVersions");
    const shaped = { ...(await workbook("old")), id: live.id, opaque: [{ name: "xl/workbook.xml", data: "PHdvcmtib29rLz4=" }] };
    const target = await recordWorkbookVersion({ docId: live.id, model: shaped, author: "You", trigger: "commit", cachedPrev: null });
    for (const blobKey of [...idb.versions.keys()].filter((key) => key.startsWith("blob:"))) {
      idb.versions.delete(blobKey);
    }
    await history.restore(target.newEntry!.id);
    expect(history.notice.value).toBe("This version can't be opened — part of it is missing");
    expect(model.value).toBe(live);
    expect(resetUndo).not.toHaveBeenCalled();
    await history.preview(target.newEntry!.id);
    expect(history.previewModel.value).toBeNull();
  });
});

describe("preview (AC-52, AC-53)", () => {
  it("shows the version in the grid on its own active sheet, read-only, and returns the live workbook untouched", async () => {
    const { history, store, live, model, previewing, show, syncView } = await harness();
    const { recordWorkbookVersion } = await import("../../lib/spreadsheetVersions");
    const version = await recordWorkbookVersion({ docId: live.id, model: { ...(await workbook("then")), id: live.id, activeSheetIndex: 0 }, author: "You", trigger: "commit", cachedPrev: null });
    const liveSheets = live.sheets;

    await history.preview(version.newEntry!.id);
    expect(syncView).toHaveBeenCalledWith(live);
    expect(previewing.value).toBe(true);
    expect(model.value).toBe(history.previewModel.value);
    expect(model.value!.activeSheetIndex).toBe(0);
    expect(model.value!.sheets[0]!.rows[0]!.cells[0]!.value).toBe("then0");
    // The store never sees the preview.
    expect(store.getStoredRecord(live.id)!.model).toBe(live);

    history.exitPreview();
    expect(previewing.value).toBe(false);
    expect(history.previewModel.value).toBeNull();
    expect(model.value).toBe(live);
    expect(live.sheets).toBe(liveSheets);
    expect(show).toHaveBeenLastCalledWith(live);
  });
});

describe("the pane and its labels (AC-48..AC-51)", () => {
  it("opens and toggles the side panel's history pane", async () => {
    const { history, sidePanel } = await harness();
    history.toggle();
    expect(sidePanel.activePane.value).toBe("history");
    expect(history.isOpen.value).toBe(true);
    history.toggle();
    expect(sidePanel.activePane.value).toBeNull();
    history.open();
    expect(sidePanel.activePane.value).toBe("history");
  });

  it("says each version's sheets, cells and cell change, and downloads in the workbook's own format", async () => {
    const { history } = await harness();
    const { cellDeltaLabel, workbookDownloadLabels } = await import("./workbookHistoryLabels");
    expect(history.statLabel({ id: "v", timestamp: 0, author: "Ada", words: 0, kind: "delta", parentId: "u", size: 1, sheets: 3, cells: 1204, cellDelta: 12 } as never)).toBe(
      "3 sheets · 1,204 cells · +12 cells",
    );
    expect(cellDeltaLabel(-3)).toBe("−3 cells");
    expect(cellDeltaLabel(0)).toBe("no change in cell count");
    expect(cellDeltaLabel(1)).toBe("+1 cell");
    expect(history.downloadLabels.value.button).toBe("Download .xlsx");
    expect(workbookDownloadLabels("ods").button).toBe("Download .ods");
    expect(workbookDownloadLabels("xlsx", "macroEnabled").button).toBe("Download .xlsm");
  });

  it("names the too-large refusal and normal exhaustion apart", async () => {
    const { history, live, emit } = await harness();
    const { WORKBOOK_HISTORY_TOO_LARGE_MESSAGE } = await import("../../lib/spreadsheetVersions");
    const { HISTORY_PAUSED_MESSAGE } = await import("../../lib/core/versionChain");
    idb.versions.set(`versions:${live.id}`, { id: `versions:${live.id}`, store: "index", docId: live.id, versions: [], paused: true, pausedReason: "too-large" } as never);
    emit("begin");
    await settle();
    expect(history.pausedMessage.value).toBe(WORKBOOK_HISTORY_TOO_LARGE_MESSAGE);
    idb.versions.set(`versions:${live.id}`, { id: `versions:${live.id}`, store: "index", docId: live.id, versions: [], paused: true } as never);
    emit("begin");
    await settle();
    expect(history.pausedMessage.value).toBe(HISTORY_PAUSED_MESSAGE);
  });

  it("in a tab that does not own the workbook every action that writes is refused", async () => {
    const { history, live, owns, emit } = await harness({ owns: false });
    expect(history.readOnly.value).toBe(true);
    edit(live, "edited");
    emit("edit");
    emit("commitPoint");
    await settle();
    expect(versionTriggers()).toEqual([]);
    owns.value = true;
    expect(history.readOnly.value).toBe(false);
  });
});
