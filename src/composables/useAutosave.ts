// The autosave state machine (SPEC-010): owns the app bar's save state, the
// dirty flag, the save cadence, recovery (Keep/Discard) and the tab-close
// behaviour. DocumentView owns everything DOM-shaped — the caret, the scroll
// container, the editable surface — and hands it in via callbacks.
import { computed, onBeforeUnmount, ref } from "vue";
import {
  createAutosaveScheduler,
  type AutosaveScheduler,
} from "../lib/core/autosaveScheduler";
import {
  commitDocument,
  deleteDraft,
  loadDraft,
  refreshDocumentFromStorage,
  replaceDocumentModel,
  saveDraft,
} from "../lib/documentStore";
import type { DocumentModel } from "../lib/documentModel";
import type { CaretPosition } from "../lib/textEditing";

/** The single failure reason SPEC-010 defines: quota and write failures both
 * surface as a full store, with a download link as the escape hatch. */
export const STORAGE_FULL_REASON = "browser storage is full";

export const SAVE_STATE_SAVING = "Saving…";
export const SAVE_STATE_SAVED = "All changes saved";
export const SAVE_STATE_ERROR = `Not saved — ${STORAGE_FULL_REASON}`;

export interface RestoredState {
  model: DocumentModel;
  caret: CaretPosition | null;
  scroll: number;
  /** True when the model is recovered draft content (the bar is showing). */
  recovered: boolean;
}

export interface UseAutosaveOptions {
  getCaret: () => CaretPosition | null;
  getScroll: () => number;
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

interface CommittedStash {
  model: DocumentModel;
  caret: CaretPosition | null;
  scroll: number;
}

export function useAutosave(options: UseAutosaveOptions) {
  const saveState = ref<"saved" | "saving" | "error">("saved");
  const saveStateText = computed(() => {
    switch (saveState.value) {
      case "saving":
        return SAVE_STATE_SAVING;
      case "error":
        return SAVE_STATE_ERROR;
      default:
        return SAVE_STATE_SAVED;
    }
  });
  /** Non-null while the recovery bar should be visible. Holds only the
   * timestamp; the committed state stays in a closure so the model never
   * passes through Vue's deep reactivity (it must stay structured-cloneable
   * and cheap). */
  const recovery = ref<{ savedAt: number } | null>(null);
  let committedStash: CommittedStash | null = null;

  let docId: string | null = null;
  let dirty = false;
  let editCount = 0;
  let scheduler: AutosaveScheduler | null = null;

  async function performSave(): Promise<void> {
    const id = docId;
    if (!id || !options.isActive()) {
      return;
    }
    const editsAtStart = editCount;
    try {
      await saveDraft(id, options.getCaret(), options.getScroll());
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
    scheduler = createAutosaveScheduler({
      flush: () => {
        void performSave();
      },
    });
    const record = await refreshDocumentFromStorage(id);
    if (docId !== id || !record) {
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
      };
      recovery.value = { savedAt: draft.savedAt };
      const model = replaceDocumentModel(id, draft.model);
      options.applyRestored({
        model,
        caret: draft.caret,
        scroll: draft.scroll,
        recovered: true,
      });
    } else {
      options.applyRestored({
        model: record.model,
        caret: record.caret ?? null,
        scroll: record.scroll ?? 0,
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
    if (commit && recovery.value === null) {
      void commitDocument(id, options.getCaret(), options.getScroll()).catch(
        () => {},
      );
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
      await commitDocument(id, options.getCaret(), options.getScroll());
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

  /** "Keep": promotes the recovered draft to committed; the bar disappears. */
  async function keepRecovered(): Promise<void> {
    const id = docId;
    if (!id || recovery.value === null) {
      return;
    }
    recovery.value = null;
    committedStash = null;
    try {
      await commitDocument(id, options.getCaret(), options.getScroll());
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
    const model = replaceDocumentModel(id, stash.model);
    dirty = false;
    // Invalidate any in-flight draft save so its completion cannot mark the
    // restored state clean against the wrong content.
    editCount += 1;
    saveState.value = "saved";
    void deleteDraft(id);
    options.applyRestored({
      model,
      caret: stash.caret,
      scroll: stash.scroll,
      recovered: false,
    });
  }

  // Tab close. beforeunload decides the leave prompt; pagehide (which also
  // fires when the prompt was skipped or confirmed) does the final write:
  // committing on a clean close, or a last-gasp draft when dirty so the next
  // open can offer recovery.
  function handleBeforeUnload(event: BeforeUnloadEvent): void {
    if (!docId || !options.isActive()) {
      return;
    }
    if (hasUnsavedChanges()) {
      scheduler?.flushNow();
      event.preventDefault();
      // Chromium still requires returnValue for the prompt to show.
      event.returnValue = "";
    } else if (recovery.value === null) {
      // Clean close: start the commit here, where more of the page
      // lifetime remains than in pagehide.
      void commitDocument(docId, options.getCaret(), options.getScroll())
        .then(() => options.onCleanClose?.())
        .catch(() => {});
    }
  }

  function handlePageHide(): void {
    if (!docId || !options.isActive()) {
      return;
    }
    if (hasUnsavedChanges()) {
      void saveDraft(docId, options.getCaret(), options.getScroll()).catch(
        () => {},
      );
    } else if (recovery.value === null) {
      void commitDocument(docId, options.getCaret(), options.getScroll())
        .then(() => options.onCleanClose?.())
        .catch(() => {});
    }
  }

  window.addEventListener("beforeunload", handleBeforeUnload);
  window.addEventListener("pagehide", handlePageHide);
  onBeforeUnmount(() => {
    window.removeEventListener("beforeunload", handleBeforeUnload);
    window.removeEventListener("pagehide", handlePageHide);
  });

  return {
    saveState,
    saveStateText,
    recovery,
    markDirty,
    hasUnsavedChanges,
    beginSession,
    endSession,
    commitNow,
    keepRecovered,
    discardRecovered,
  };
}
