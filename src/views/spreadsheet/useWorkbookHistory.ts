// The workbook's version history controller (SPEC-138) — SPEC-040's pane for
// workbooks, the controller SPEC-103's panel contract expects: the pane's
// open state (SPEC-129's sheet side panel, `"history"` — its own persisted
// state is the only record of whether the pane is open), the snapshot
// triggers (createVersionCadenceTimer's five minutes of editing activity, the
// commit points, and the pair bracketing a restore), and the panel's actions.
// Storage and the chain live in lib/spreadsheetVersions.ts; the session it
// records against is useWorkbookSaving's, which it subscribes to.
//
// Preview shows the reconstructed workbook in the grid, read-only, on its own
// sheet tabs and its own active sheet: the live workbook is set aside (its
// view written into it first) and shown again on exit — its selection and
// undo stack untouched. While previewing, no edit, view-state write, draft of
// the preview or version of it can happen (useWorkbookSaving.previewing).
//
// Restore is a REPLACE, never an undoable step: useSpreadsheetUndo's
// snapshot holds the sheets and WORKBOOK_SNAPSHOT_FIELDS — not the title, the
// theme, the properties — so an undo of a restore would leave a workbook that
// never existed. The pre-restore version is written first; the toast's "Undo
// restore" restores it the same way. Nothing here is awaited on an edit's
// path, and every failure stays inside history: the workbook's own records
// are never touched by it.
import { computed, onBeforeUnmount, ref, shallowRef, type ShallowRef } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import { useToasts } from "../../composables/useToasts";
import { createVersionCadenceTimer } from "../../lib/core/versionCadenceScheduler";
import { emptyVersionsIndex, HISTORY_PAUSED_MESSAGE, type SnapshotTrigger, type VersionIndexEntry, type VersionsIndexRecord } from "../../lib/core/versionChain";
import { replaceStoredModel } from "../../lib/documentStore";
import { downloadSpreadsheet } from "../../lib/spreadsheet/download";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import {
  deleteWorkbookVersion,
  loadWorkbookIndex,
  nameWorkbookVersion,
  reconstructWorkbookVersion,
  recordWorkbookVersion,
  removeWorkbookVersionName,
  workbookVersionEntries,
  WORKBOOK_HISTORY_TOO_LARGE_MESSAGE,
  type WorkbookVersionTail,
} from "../../lib/spreadsheetVersions";
import type { SheetSidePanel } from "./useSheetSidePanel";
import type { WorkbookSaving } from "./useWorkbookSaving";
import {
  UNDO_RESTORE_LABEL,
  VERSION_MISSING_MESSAGE,
  VERSION_RESTORED_MESSAGE,
  workbookDownloadLabels,
  workbookStatLabel,
} from "./workbookHistoryLabels";

export interface WorkbookHistoryDeps {
  /** useWorkbookSaving's: the session, who owns it, the commit, and the
   * flag that makes the view read-only while a version is previewed. */
  saving: Pick<WorkbookSaving, "listen" | "owns" | "previewing" | "commitNow">;
  /** The model on show (useWorkbookLoad's `workbook`). */
  model: ShallowRef<SpreadsheetModel | null>;
  /** Shows another model object in place of the one on show. */
  show: (model: SpreadsheetModel) => void;
  /** Writes the view (each sheet's selection, the sheet on show) into the
   * live model before a preview sets it aside (useSheetViewState.syncInto). */
  syncView: (model: SpreadsheetModel) => void;
  /** useSpreadsheetUndo.reset — a restore clears the stack. */
  resetUndo: () => void;
  /** SPEC-129's side panel — the pane is its `"history"`. */
  sidePanel: Pick<SheetSidePanel, "open" | "toggle" | "activePane">;
}

/** SPEC-116's reader-version upgrade, on a reconstructed version (it may
 * predate a reader bump) — loaded on demand, as the store loads it. */
async function upgraded(model: SpreadsheetModel): Promise<SpreadsheetModel> {
  const [{ upgradeSpreadsheetModel }, { readSpreadsheetOffThread }] = await Promise.all([
    import("../../lib/spreadsheet/open"),
    import("../../lib/workers/spreadsheetParseClient"),
  ]);
  return upgradeSpreadsheetModel(model, readSpreadsheetOffThread);
}

export function useWorkbookHistory(deps: WorkbookHistoryDeps) {
  const { preferences } = usePreferences();
  const { addToast } = useToasts();

  // --- the pane -------------------------------------------------------------------
  const isOpen = computed(() => deps.sidePanel.activePane.value === "history");
  function open(): void {
    deps.sidePanel.open("history");
  }
  function toggle(): void {
    deps.sidePanel.toggle("history");
  }

  const index = shallowRef<VersionsIndexRecord>(emptyVersionsIndex(""));
  const versions = computed(() => workbookVersionEntries(index.value));
  const namedOnly = ref(false);
  const visibleVersions = computed(() => (namedOnly.value ? versions.value.filter((v) => !!v.name) : versions.value));
  const pausedMessage = computed(() => {
    if (!index.value.paused) {
      return "";
    }
    return index.value.pausedReason === "too-large" ? WORKBOOK_HISTORY_TOO_LARGE_MESSAGE : HISTORY_PAUSED_MESSAGE;
  });
  /** The panel's last word on an action (a version that can't be opened). */
  const notice = ref("");
  /** Another tab holds the workbook: no action here writes. */
  const readOnly = computed(() => !deps.saving.owns.value);

  // --- preview: the version in the grid, read-only --------------------------------
  const previewModel = shallowRef<SpreadsheetModel | null>(null);
  const previewVersionId = ref<string | null>(null);
  /** The live workbook while a preview is on show. */
  let setAside: SpreadsheetModel | null = null;
  /** The workbook the session edits — never the preview. */
  const live = (): SpreadsheetModel | null => setAside ?? deps.model.value;

  /** The version's model, upgraded, or null — and the panel says why. */
  async function versionModel(id: string, versionId: string): Promise<SpreadsheetModel | null> {
    const version = await reconstructWorkbookVersion(id, versionId, live() ?? undefined);
    if (!version) {
      notice.value = VERSION_MISSING_MESSAGE;
      return null;
    }
    try {
      return await upgraded(version);
    } catch {
      notice.value = VERSION_MISSING_MESSAGE;
      return null;
    }
  }

  async function preview(versionId: string): Promise<void> {
    const id = docId;
    if (!id || !live()) {
      return;
    }
    notice.value = "";
    const model = await versionModel(id, versionId);
    const current = live();
    if (docId !== id || !model || !current) {
      return;
    }
    if (!setAside) {
      deps.syncView(current);
      setAside = current;
    }
    previewModel.value = model;
    previewVersionId.value = versionId;
    deps.saving.previewing.value = true;
    deps.show(model);
  }

  /** Ends a preview: the live workbook back on show — unless the view has
   * moved on (another workbook is on show), which it must not cover. */
  function exitPreview(): void {
    const restored = deps.model.value === previewModel.value ? setAside : null;
    setAside = null;
    previewModel.value = null;
    previewVersionId.value = null;
    deps.saving.previewing.value = false;
    if (restored) {
      deps.show(restored);
    }
  }


  // --- the session -------------------------------------------------------------------
  let docId: string | null = null;
  let cachedPrev: WorkbookVersionTail | null = null;
  let changedSinceLastVersion = false;

  /** Snapshots run one at a time: each appends to the index the one before
   * it wrote (two at once would each append to the same index, and the later
   * write would drop the earlier version and orphan its blob references). */
  let snapshots: Promise<unknown> = Promise.resolve();
  function snapshot(trigger: SnapshotTrigger): Promise<string | null> {
    const next = snapshots.then(() => recordSnapshot(trigger));
    snapshots = next.catch(() => null);
    return next;
  }

  async function recordSnapshot(trigger: SnapshotTrigger): Promise<string | null> {
    const id = docId;
    const model = live();
    if (!id || !model || !deps.saving.owns.value) {
      return null;
    }
    // "Keep version history" off stops the periodic and commit-point
    // recording only; a restore's own pair still runs, so undoing a restore
    // can never lose the pre-restore state.
    if (!preferences.keepVersionHistory && (trigger === "cadence" || trigger === "commit")) {
      return null;
    }
    const result = await recordWorkbookVersion({ docId: id, model, author: preferences.authorName, trigger, cachedPrev });
    if (docId !== id) {
      return null;
    }
    index.value = result.index;
    if (!result.newEntry) {
      return null;
    }
    cachedPrev = result.newEntry;
    changedSinceLastVersion = false;
    return result.newEntry.id;
  }

  const cadence = createVersionCadenceTimer({
    onDue: () => {
      if (changedSinceLastVersion) {
        void snapshot("cadence");
      }
    },
  });
  onBeforeUnmount(() => cadence.dispose());

  deps.saving.listen({
    begin: (id) => {
      docId = id;
      cachedPrev = null;
      changedSinceLastVersion = false;
      namedOnly.value = false;
      notice.value = "";
      const empty = emptyVersionsIndex(id);
      index.value = empty;
      void loadWorkbookIndex(id).then((loaded) => {
        // A snapshot or a panel action that landed first holds the newer index.
        if (docId === id && index.value === empty) {
          index.value = loaded;
        }
      });
    },
    end: () => {
      cadence.reset();
      exitPreview();
      docId = null;
      cachedPrev = null;
      changedSinceLastVersion = false;
    },
    // Five minutes of editing, only when the workbook changed.
    edit: () => {
      changedSinceLastVersion = true;
      cadence.noteChange();
    },
    // A download, a clean close: a version when the workbook changed since
    // its last one.
    commitPoint: () => {
      if (changedSinceLastVersion) {
        void snapshot("commit");
      }
    },
    // The session shows another workbook in place of the one on show (an
    // adopted commit, a recovered draft, Discard): a preview ends first —
    // leaving it later would put back what was replaced.
    swap: () => exitPreview(),
  });

  // --- restore: a replace -------------------------------------------------------
  /** pre-restore snapshot → reconstruct → upgrade → replaceStoredModel →
   * show → reset the undo stack → commit → post-restore snapshot → the
   * toast, whose "Undo restore" restores the pre-restore version the same
   * way. */
  async function restore(versionId: string): Promise<void> {
    exitPreview();
    const id = docId;
    if (!id || !deps.model.value || readOnly.value) {
      return;
    }
    notice.value = "";
    const before = await snapshot("pre-restore");
    const target = await versionModel(id, versionId);
    if (docId !== id || !target || !deps.model.value) {
      return;
    }
    const model = replaceStoredModel(id, target) as SpreadsheetModel;
    deps.show(model);
    deps.resetUndo();
    await deps.saving.commitNow();
    await snapshot("post-restore");
    const undoTarget = before;
    addToast(VERSION_RESTORED_MESSAGE, undoTarget ? { label: UNDO_RESTORE_LABEL, run: () => void restore(undoTarget) } : undefined);
  }

  /** The toast's action, for callers holding no toast: restores the newest
   * pre-restore version. */
  function undoRestore(): Promise<void> {
    const before = versions.value.find((entry) => entry.trigger === "pre-restore");
    return before ? restore(before.id) : Promise.resolve();
  }

  // --- name / unname / delete / download ---------------------------------------
  /** Runs one index action, adopting the index it returns; a failure is
   * said in the panel and leaves the list as it was. */
  async function update(action: (id: string, model: SpreadsheetModel) => Promise<VersionsIndexRecord>): Promise<void> {
    const id = docId;
    const model = live();
    if (!id || !model || readOnly.value) {
      return;
    }
    try {
      const next = await action(id, model);
      if (docId === id) {
        index.value = next;
      }
    } catch {
      notice.value = "Version history couldn't be updated";
    }
  }

  function nameVersion(versionId: string, name: string): Promise<void> {
    return update((id) => nameWorkbookVersion(id, versionId, name));
  }

  function removeName(versionId: string): Promise<void> {
    return update((id, model) => removeWorkbookVersionName(id, versionId, model));
  }

  function deleteVersion(versionId: string): Promise<void> {
    if (previewVersionId.value === versionId) {
      exitPreview();
    }
    return update((id, model) => deleteWorkbookVersion(id, versionId, model));
  }

  /** Downloads a version in the workbook's own format and variant. */
  async function downloadVersion(versionId: string): Promise<void> {
    const id = docId;
    if (!id) {
      return;
    }
    notice.value = "";
    const model = await versionModel(id, versionId);
    if (model) {
      await downloadSpreadsheet(model);
    }
  }

  const downloadLabels = computed(() => {
    const model = live();
    return workbookDownloadLabels(model?.format ?? "xlsx", model?.packageVariant);
  });

  return {
    open,
    toggle,
    isOpen,
    versions,
    visibleVersions,
    namedOnly,
    pausedMessage,
    notice,
    readOnly,
    previewModel,
    previewVersionId,
    preview,
    exitPreview,
    restore,
    undoRestore,
    nameVersion,
    removeName,
    deleteVersion,
    downloadVersion,
    statLabel: (entry: VersionIndexEntry) => workbookStatLabel(entry),
    downloadLabels,
  };
}

export type WorkbookHistory = ReturnType<typeof useWorkbookHistory>;

/** Narrows a component's runtime-typed `history` prop (the .vue parser bars
 * casts). */
export function asWorkbookHistory(value: unknown): WorkbookHistory {
  return value as WorkbookHistory;
}
