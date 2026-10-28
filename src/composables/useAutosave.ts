// The autosave state machine (SPEC-010): owns the app bar's save state, the
// dirty flag, the save cadence, recovery (Keep/Discard) and the tab-close
// behaviour. The view owns everything DOM-shaped — the caret, the scroll
// container, the editable surface — and hands it in via callbacks. It serves
// every editing product: DocumentView, the deck's useDeckSaving (SPEC-103)
// and the workbook's useWorkbookSaving (SPEC-138).
import { computed, ref } from "vue";
import {
  createAutosaveScheduler,
  type AutosaveScheduler,
} from "../lib/core/autosaveScheduler";
import { usePreferences } from "./usePreferences";
import { useTabClose } from "./useTabClose";
import {
  commitDocument,
  deleteDraft,
  isMediaModel,
  isSpreadsheetModel,
  isSpreadsheetResident,
  loadDraft,
  loadSpreadsheet,
  refreshDocumentFromStorage,
  replaceStoredModel,
  saveDraft,
  type DraftState,
  type StoredModel,
} from "../lib/documentStore";
import type { SheetViewState } from "../lib/spreadsheet/selection";
import { saveStateText as textForSaveState, STORAGE_FULL_REASON, SAVE_STATE_SAVING, SAVE_STATE_SAVED, SAVE_STATE_ERROR, type SaveState } from "../lib/core/saveState";
import { recoveryNotice } from "../lib/format";
import { now } from "../lib/clock";
import type { CaretPosition } from "../lib/textEditing";

// The state vocabulary is product-neutral (the deck reports through it too),
// so it lives in lib/core/saveState.ts. Re-exported for the callers that had
// been importing it from here.
export { STORAGE_FULL_REASON, SAVE_STATE_SAVING, SAVE_STATE_SAVED, SAVE_STATE_ERROR };

export interface RestoredState {
  /** A text document's model, or a deck's (SPEC-103). */
  model: StoredModel;
  caret: CaretPosition | null;
  scroll: number;
  /** A deck's selected slide (SPEC-103); absent for text. */
  slideIndex?: number;
  /** A workbook's tab, active cell and top-left cell (SPEC-138). */
  sheetView?: SheetViewState;
  /** True when the model is recovered draft content (the bar is showing). */
  recovered: boolean;
}

export interface UseAutosaveOptions {
  getCaret: () => CaretPosition | null;
  getScroll: () => number;
  /** A deck's selected slide, recorded with each draft and commit (SPEC-103). */
  getSlideIndex?: () => number;
  /** A workbook's view, recorded with each draft and commit (SPEC-138). */
  getSheetView?: () => SheetViewState | undefined;
  /** Whether this tab currently owns the document (edit lock held and a
   * document open). While false the composable never writes. */
  isActive: () => boolean;
  /** Applies a restored model + caret/scroll to the editor: on session
   * start, when recovered draft content is loaded, and on "Discard". */
  applyRestored: (state: RestoredState) => void;
  /** Called after a clean tab-close commit succeeds — SPEC-040's other
   * commit-point version-history trigger, alongside commitNow's download. */
  onCleanClose?: () => void;
}

/** The committed state "Discard" restores (the store's replaceStoredModel
 * swaps it back through its own kind's path, SPEC-138). */
type CommittedStash = Omit<RestoredState, "recovered">;

export function useAutosave(options: UseAutosaveOptions) {
  // SPEC-066: crash-recovery drafts (the debounced saveDraft below) are the
  // safety net and always run; the "Autosave" preference only gates the
  // automatic *committed* writes at clean-close/pagehide, so turning it off
  // stops the committed write, not the safety net (AC-47). The "Autosave
  // delay" preference feeds the scheduler's debounce, read live per change.
  const { preferences } = usePreferences();
  const saveState = ref<SaveState>("saved");
  /** Non-null while the recovery bar should be visible. Holds only the
   * timestamp; the committed state stays in a closure so the model never
   * passes through Vue's deep reactivity (it must stay structured-cloneable
   * and cheap). */
  const recovery = ref<{ savedAt: number } | null>(null);
  /** What the bar shows. A pending recovery decision outranks the stored state:
   * `commitNow` and `endSession` both stand down while the banner is up, so the
   * bar claiming "All changes saved" beside a banner announcing recovered
   * UNSAVED changes was two contradictory statements about one document. */
  const displaySaveState = computed<SaveState>(() => (recovery.value ? "recovered" : saveState.value));
  const saveStateText = computed(() => textForSaveState(displaySaveState.value));
  /** The banner's sentence: relative, and off the injectable clock. */
  const recoveryLabel = computed(() => (recovery.value ? recoveryNotice(recovery.value.savedAt, now().getTime()) : ""));
  let committedStash: CommittedStash | null = null;

  let docId: string | null = null;
  let dirty = false;
  let editCount = 0;
  let scheduler: AutosaveScheduler | null = null;

  /** The editor state every draft and commit records beside the model. */
  const draftState = (): DraftState => {
    const slide = options.getSlideIndex ? { slideIndex: options.getSlideIndex() } : {};
    const sheetView = options.getSheetView?.();
    return { caret: options.getCaret(), scroll: options.getScroll(), ...slide, ...(sheetView ? { sheetView } : {}) };
  };

  function makeScheduler(): AutosaveScheduler {
    return createAutosaveScheduler({
      flush: () => {
        void performSave();
      },
      debounceMs: () => preferences.autosaveDelayMs,
    });
  }

  async function performSave(): Promise<void> {
    const id = docId;
    if (!id || !options.isActive()) {
      return;
    }
    const editsAtStart = editCount;
    try {
      await saveDraft(id, draftState());
      if (docId !== id) {
        return;
      }
      // Only a save that covers every edit so far makes the document clean;
      // edits made while the write was in flight keep it dirty and the
      // scheduler already has their debounce pending.
      if (editCount === editsAtStart) {
        dirty = false;
        saveState.value = "saved";
      }
      scheduler?.notifySaveResult(true);
    } catch {
      if (docId !== id) {
        return;
      }
      saveState.value = "error";
      scheduler?.notifySaveResult(false);
    }
  }

  /** Report a content edit: arms the save cadence and the save state. */
  function markDirty(): void {
    if (!docId) {
      return;
    }
    dirty = true;
    editCount += 1;
    // A failed state stays visible until a write actually succeeds
    // (SPEC-010: the failure clears on success, not on further typing).
    if (saveState.value !== "error") {
      saveState.value = "saving";
    }
    scheduler?.noteChange();
  }

  /** True while leaving would lose data: edits not yet fully written (a
   * write pending or in flight) or a failed write. Deliberately narrow —
   * see the spec's "Leave-confirmation" note. */
  function hasUnsavedChanges(): boolean {
    return dirty || saveState.value === "error";
  }

  /** Starts the autosave session for a document this tab now owns: adopts a
   * newer committed record from another tab if there is one, then checks
   * for a crash draft and enters recovery when the draft is newer. */
  async function beginSession(id: string): Promise<void> {
    docId = id;
    dirty = false;
    editCount = 0;
    saveState.value = "saved";
    recovery.value = null;
    committedStash = null;
    scheduler?.dispose();
    scheduler = makeScheduler();
    const record = await refreshDocumentFromStorage(id);
    if (record && isSpreadsheetModel(record.model) && !isSpreadsheetResident(id)) {
      // SPEC-138: another tab's commit is adopted as its shell — rejoin the
      // rows and sources before the session hands the view the workbook.
      await loadSpreadsheet(id);
    }
    if (docId !== id || !record) {
      return;
    }
    if (isMediaModel(record.model)) {
      // A media id can only reach here through a misrouted view, and a
      // session would clobber the stored bytes (SPEC-107).
      return;
    }
    const draft = await loadDraft(id);
    if (docId !== id) {
      return;
    }
    if (draft && draft.savedAt > record.lastModified) {
      // The previous session ended without a clean close: show the draft
      // content with the recovery bar, keeping the committed state aside
      // for "Discard".
      committedStash = {
        model: record.model,
        caret: record.caret ?? null,
        scroll: record.scroll ?? 0,
        slideIndex: record.slideIndex,
        sheetView: record.sheetView,
      };
      recovery.value = { savedAt: draft.savedAt };
      const model = replaceStoredModel(id, draft.model);
      options.applyRestored({
        model,
        caret: draft.caret,
        scroll: draft.scroll,
        slideIndex: draft.slideIndex,
        sheetView: draft.sheetView,
        recovered: true,
      });
    } else {
      options.applyRestored({
        model: record.model,
        caret: record.caret ?? null,
        scroll: record.scroll ?? 0,
        slideIndex: record.slideIndex,
        sheetView: record.sheetView,
        recovered: false,
      });
    }
  }

  /** Ends the session. An in-app close (navigation, unmount) is a clean
   * close, so the final state is committed — unless the recovery bar is
   * still open, in which case the draft must survive for the next open to
   * offer recovery again. */
  function endSession(commit: boolean): void {
    const id = docId;
    if (!id) {
      return;
    }
    scheduler?.dispose();
    scheduler = null;
    // The committed write is gated on the Autosave preference (AC-47); the
    // draft already written by the scheduler remains the recoverable copy.
    if (commit && recovery.value === null && preferences.autosave) {
      void commitDocument(id, draftState()).catch(() => {});
    }
    docId = null;
    dirty = false;
    recovery.value = null;
    committedStash = null;
    saveState.value = "saved";
  }

  /** Commits the current state right now — the "explicit save" commit point
   * for downloads (SPEC-006/007). Skipped while the recovery bar is open so
   * downloading a recovered document does not silently decide "Keep". */
  async function commitNow(): Promise<void> {
    const id = docId;
    if (!id || !options.isActive() || recovery.value !== null) {
      return;
    }
    const editsAtStart = editCount;
    try {
      await commitDocument(id, draftState());
      if (docId !== id) {
        return;
      }
      if (editCount === editsAtStart) {
        dirty = false;
        saveState.value = "saved";
      }
    } catch {
      if (docId !== id) {
        return;
      }
      saveState.value = "error";
      scheduler?.notifySaveResult(false);
    }
  }

  /** The close path's flush (SPEC-060): cancels the pending debounce and
   * runs — awaited — the committed write it was deferring, so the library
   * reads the just-saved content, never a stale row. Returns whether it is
   * safe to leave: true on a clean commit (or when there is nothing to
   * flush), false on a write failure so the caller can stay and toast. The
   * recovery bar is the one case that leaves without committing: keeping the
   * draft so reopening still offers recovery (AC-26). While in flight the
   * save indicator reads "Saving…" (AC-10). */
  async function flushForClose(): Promise<boolean> {
    const id = docId;
    if (!id || !options.isActive()) {
      return true;
    }
    if (recovery.value !== null) {
      return true;
    }
    const editsAtStart = editCount;
    saveState.value = "saving";
    // Cancel the pending debounce; a fresh scheduler stands ready in case the
    // write fails and the session stays open (its edits must still autosave).
    scheduler?.dispose();
    scheduler = makeScheduler();
    try {
      await commitDocument(id, draftState());
      if (docId !== id) {
        return true;
      }
      if (editCount === editsAtStart) {
        dirty = false;
        saveState.value = "saved";
      }
      return true;
    } catch {
      if (docId !== id) {
        return false;
      }
      saveState.value = "error";
      scheduler?.notifySaveResult(false);
      return false;
    }
  }

  /** "Keep": promotes the recovered draft to committed; the bar disappears. */
  async function keepRecovered(): Promise<void> {
    const id = docId;
    if (!id || recovery.value === null) {
      return;
    }
    recovery.value = null;
    committedStash = null;
    try {
      await commitDocument(id, draftState());
      if (docId !== id) {
        return;
      }
      dirty = false;
      saveState.value = "saved";
    } catch {
      if (docId !== id) {
        return;
      }
      saveState.value = "error";
      scheduler?.notifySaveResult(false);
    }
  }

  /** "Discard": deletes the draft and restores the committed record — the
   * last clean state; the bar disappears. */
  function discardRecovered(): void {
    const id = docId;
    if (!id || recovery.value === null || !committedStash) {
      return;
    }
    const stash = committedStash;
    recovery.value = null;
    committedStash = null;
    const model = replaceStoredModel(id, stash.model);
    dirty = false;
    // Invalidate any in-flight draft save so its completion cannot mark the
    // restored state clean against the wrong content.
    editCount += 1;
    saveState.value = "saved";
    void deleteDraft(id);
    options.applyRestored({ ...stash, model, recovered: false });
  }

  // Tab close — the leave prompt, the clean-close commit and the last-gasp
  // draft — lives in useTabClose, reading this session's state.
  useTabClose({
    sessionId: () => docId,
    isActive: options.isActive,
    hasUnsavedChanges,
    flushNow: () => scheduler?.flushNow(),
    mayCommit: () => recovery.value === null && preferences.autosave,
    commit: (id) => commitDocument(id, draftState()),
    saveDraft: (id) => saveDraft(id, draftState()),
    onCleanClose: () => options.onCleanClose?.(),
  });

  return {
    saveState,
    displaySaveState,
    saveStateText,
    recovery,
    recoveryLabel,
    markDirty,
    hasUnsavedChanges,
    beginSession,
    endSession,
    commitNow,
    flushForClose,
    keepRecovered,
    discardRecovered,
  };
}
