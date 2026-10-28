// The workbook's saving (SPEC-138): SPEC-010's autosave, recovery and edit
// lock, for the workbook view. One state machine, not a fourth — useAutosave
// serves workbooks as it serves text documents and decks: drafts on its 2 s /
// 30 s cadence (the store writes them as draft chunk records beside the
// committed ones), commits at SPEC-010's four points (create/import — the
// store's own; a download; a clean close — Close through flushForClose, any
// other route away, the tab closing; and "Keep"), recovery when the draft is
// newer than the commit, with the workbook's state `{ caret: null, scroll: 0,
// sheetView }`. SpreadsheetView makes one `useWorkbookSaving({ … })` call and
// hooks the template up; every edit reaches here through SPEC-121's edit
// pipeline — its `onApplied` calls noteEdit() where it wrote the workbook
// through (touchDocument) before.
//
// Two tabs editing one workbook would race on its row chunk records, so the
// workbook takes SPEC-010's edit lock (core/editLock.ts) on open: while
// another tab holds it (`busy`) this one is read-only — no session, no draft,
// no commit, no version — and when that tab closes the queued lock grants,
// the session adopts its commit (rows rejoined), the undo stack resets and
// editing enables, with no reload. The lock goes on Close and on unmount.
//
// useWorkbookHistory (the version history controller) subscribes to the
// session here — its begin and end, every edit, every commit point.
import { computed, inject, onBeforeUnmount, ref, watch, type Ref, type ShallowRef } from "vue";
import { matchedRouteKey, onBeforeRouteLeave } from "vue-router";
import { useAutosave } from "../../composables/useAutosave";
import { useToasts } from "../../composables/useToasts";
import { acquireEditLock, type EditLockHandle, type LockManagerLike } from "../../lib/core/editLock";
import { STORAGE_FULL_REASON } from "../../lib/core/saveState";
import { isSpreadsheetModel, replaceSpreadsheetModel } from "../../lib/documentStore";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetViewState } from "../../lib/spreadsheet/selection";
import { withSheetView } from "./useSheetViewState";

/** The conflict bar's words (SPEC-010's two-tab rule, for workbooks). */
export const WORKBOOK_CONFLICT_MESSAGE = "This workbook is open in another tab";

export type WorkbookLockState = "pending" | "acquired" | "busy";

export interface WorkbookSavingDeps {
  /** The model on show (useWorkbookLoad's `workbook`). */
  model: ShallowRef<SpreadsheetModel | null>;
  workbookId: Readonly<Ref<string>>;
  /** Where the view stands — recorded with each draft and commit. */
  getSheetView: () => SheetViewState | undefined;
  /** Puts a recorded view back on the model on show; false when it names
   * a sheet that no longer exists (it is then ignored). */
  applySheetView: (view: SheetViewState | undefined) => boolean | void;
  /** useSpreadsheetUndo.reset — the stack's by-reference snapshots point at
   * rows a swapped-in workbook no longer has. */
  resetUndo: () => void;
  /** The workbook download (the save badge's error link). */
  download: () => void;
  /** Shows another model object in place of the one on show
   * (useWorkbookLoad.replace) — a recovered or discarded workbook. */
  show: (model: SpreadsheetModel) => void;
  /** Re-mounts the grid over the same model, so it scrolls to a restored
   * top-left cell (useWorkbookLoad.remount). */
  remount?: () => void;
  /** SPEC-137: the workbook's storage came back incomplete — read-only, so
   * no session ever starts. */
  storageIncomplete?: Readonly<Ref<boolean>>;
  /** The Web Locks seam (tests inject one; core/editLock's default is the
   * browser's). */
  locks?: LockManagerLike;
}

/** What useWorkbookHistory hears from the session. */
export interface WorkbookSessionListener {
  /** A workbook was opened in this view (whichever tab holds its lock). */
  begin?: (id: string) => void;
  end?: () => void;
  /** An edit was noted (useSpreadsheetUndo's pipeline). */
  edit?: () => void;
  /** A commit point: a download, a clean close. */
  commitPoint?: () => void;
}

export function useWorkbookSaving(deps: WorkbookSavingDeps) {
  const { addToast } = useToasts();
  /** "pending" until the Web Lock resolves; editing waits for "acquired". */
  const lockState = ref<WorkbookLockState>("pending");
  /** Set by useWorkbookHistory while a version is previewed in the grid. */
  const previewing = ref(false);
  /** This tab owns the workbook: it may draft, commit and record versions. */
  const owns = computed(() => lockState.value === "acquired");
  /** Edits are allowed: owned, not previewing, storage complete. */
  const editable = computed(() => owns.value && !previewing.value && deps.storageIncomplete?.value !== true);
  const listeners: WorkbookSessionListener[] = [];

  function emit(event: keyof WorkbookSessionListener, id = ""): void {
    for (const listener of listeners) {
      listener[event]?.(id);
    }
  }

  /** Shows a restored workbook with its recorded view: a new model object
   * through `show` (the grid re-mounts on it), the same one re-mounted when
   * the view changed what shows. */
  function applyWorkbook(model: SpreadsheetModel, view: SheetViewState | undefined): void {
    if (model !== deps.model.value) {
      withSheetView(model, view);
      deps.show(model);
    } else if (deps.applySheetView(view)) {
      deps.remount?.();
    }
  }

  const autosave = useAutosave({
    getCaret: () => null,
    getScroll: () => 0,
    getSheetView: deps.getSheetView,
    isActive: () => owns.value && deps.model.value !== null,
    applyRestored: (state) => {
      if (isSpreadsheetModel(state.model)) {
        applyWorkbook(state.model, state.sheetView);
      }
    },
    // SPEC-040's commit-point trigger for a clean tab close.
    onCleanClose: () => emit("commitPoint"),
  });

  // --- the session and the lock ------------------------------------------------
  // The session follows the routed workbook: it opens once that workbook is
  // on show and ends — a clean close, so it commits when this tab owns it —
  // when the route moves on or the view unmounts. Swapping the model of the
  // same workbook (recovery, a restore, a preview) is not a new session.
  let sessionId: string | null = null;
  let lock: EditLockHandle | null = null;

  function open(id: string): void {
    sessionId = id;
    lockState.value = "pending";
    emit("begin", id);
    lock = acquireEditLock(
      id,
      (state) => {
        if (sessionId !== id) {
          return;
        }
        const waited = lockState.value === "busy";
        lockState.value = state;
        if (state === "acquired") {
          if (waited) {
            // A tab that waited behind another adopts its commit: the
            // stack's snapshots belong to the workbook it loaded first. (A
            // first acquisition has nothing stacked — editing waits for it.)
            deps.resetUndo();
          }
          void autosave.beginSession(id);
        }
      },
      deps.locks,
    );
  }

  function close(): void {
    if (sessionId === null) {
      return;
    }
    autosave.endSession(owns.value);
    emit("end");
    lock?.release();
    lock = null;
    sessionId = null;
    lockState.value = "pending";
  }

  watch(
    () => {
      const model = deps.model.value;
      return model && model.id === deps.workbookId.value && deps.storageIncomplete?.value !== true ? model.id : null;
    },
    (id) => {
      if (id === sessionId) {
        return;
      }
      close();
      if (id !== null) {
        open(id);
      }
    },
    { immediate: true },
  );
  onBeforeUnmount(close);

  // --- edits and commit points -------------------------------------------------

  /** Every finished edit, undo and redo (SPEC-121's onApplied): the live
   * model becomes the store's — uncommitted — and the draft cadence and the
   * version cadence arm. Opening, switching tabs, moving the selection and
   * the open-time recalculation never come through here. */
  function noteEdit(): void {
    const model = deps.model.value;
    if (!model || !editable.value) {
      return;
    }
    replaceSpreadsheetModel(model.id, model);
    autosave.markDirty();
    emit("edit");
  }

  /** A download — SPEC-010's explicit save: commit, then SPEC-040's
   * commit-point snapshot. */
  function commitPoint(): void {
    if (owns.value) {
      void autosave.commitNow().then(() => emit("commitPoint"));
    }
  }

  /** The close path's flush (SPEC-060): commits — awaited — before the view
   * goes, so the release that follows finds storage holding the workbook;
   * false (a failed write) keeps the workbook open. */
  async function flushForClose(): Promise<boolean> {
    const flushed = await autosave.flushForClose();
    if (flushed && owns.value) {
      emit("commitPoint");
    }
    return flushed;
  }

  // Every route away from the workbook — Close, File ▸ Close, the home link,
  // Back — is the same commit boundary (useFileActions' rule for text).
  if (inject(matchedRouteKey, null)) {
    onBeforeRouteLeave(async () => {
      const flushed = await flushForClose();
      if (!flushed) {
        addToast(
          `Couldn't close "${deps.model.value?.title ?? "this workbook"}" — ${STORAGE_FULL_REASON}. ` +
            "Your workbook is still open; retry, or download a copy.",
        );
      }
      return flushed;
    });
  }

  /** "Discard": the committed workbook back with its view, and an empty undo
   * stack. */
  function discardRecovered(): void {
    autosave.discardRecovered();
    deps.resetUndo();
  }

  return {
    saveState: autosave.displaySaveState,
    saveStateText: autosave.saveStateText,
    /** The badge shows only while this tab owns the workbook. */
    showSaveState: owns,
    recovery: autosave.recovery,
    recoveryLabel: autosave.recoveryLabel,
    keepRecovered: autosave.keepRecovered,
    discardRecovered,
    noteEdit,
    commitPoint,
    commitNow: autosave.commitNow,
    flushForClose,
    lockState,
    owns,
    editable,
    previewing,
    conflictMessage: computed(() => (lockState.value === "busy" ? WORKBOOK_CONFLICT_MESSAGE : "")),
    download: () => deps.download(),
    /** useWorkbookHistory's subscription to the session. */
    listen(listener: WorkbookSessionListener): void {
      listeners.push(listener);
      if (sessionId !== null) {
        listener.begin?.(sessionId);
      }
    },
  };
}

export type WorkbookSaving = ReturnType<typeof useWorkbookSaving>;

/** Narrows a component's runtime-typed `saving` prop (the .vue parser bars
 * casts). */
export function asWorkbookSaving(value: unknown): WorkbookSaving {
  return value as WorkbookSaving;
}
