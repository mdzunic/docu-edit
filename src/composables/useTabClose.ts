// The tab-close half of the autosave state machine (SPEC-010), split out of
// useAutosave when SPEC-103 widened it to decks and its composable budget
// (composableSize.spec.ts) asked for an extraction rather than more room. The
// behaviour is useAutosave's, moved whole: beforeunload decides the leave
// prompt; pagehide (which also fires when the prompt was skipped or
// confirmed) does the final write — committing on a clean close, or a
// last-gasp draft when dirty so the next open can offer recovery. The
// session's state stays in useAutosave and is read through these callbacks.
import { onBeforeUnmount } from "vue";

export interface TabCloseDeps {
  /** The open session's document id; null when there is none. */
  sessionId: () => string | null;
  /** Whether this tab owns the document (UseAutosaveOptions.isActive). */
  isActive: () => boolean;
  hasUnsavedChanges: () => boolean;
  /** Runs the scheduler's pending draft write now. */
  flushNow: () => void;
  /** Whether a clean close may commit: no recovery decision is pending and
   * the Autosave preference is on (SPEC-066 AC-47). */
  mayCommit: () => boolean;
  commit: (id: string) => Promise<void>;
  saveDraft: (id: string) => Promise<void>;
  /** After a clean tab-close commit succeeds (SPEC-040's commit point). */
  onCleanClose: () => void;
}

export function useTabClose(deps: TabCloseDeps): void {
  function handleBeforeUnload(event: BeforeUnloadEvent): void {
    const id = deps.sessionId();
    if (!id || !deps.isActive()) {
      return;
    }
    if (deps.hasUnsavedChanges()) {
      deps.flushNow();
      event.preventDefault();
      // Chromium still requires returnValue for the prompt to show.
      event.returnValue = "";
    } else if (deps.mayCommit()) {
      // Clean close: start the commit here, where more of the page
      // lifetime remains than in pagehide. Gated on the Autosave preference
      // (AC-47) — a pending draft still flushed above keeps data recoverable.
      void deps
        .commit(id)
        .then(() => deps.onCleanClose())
        .catch(() => {});
    }
  }

  function handlePageHide(): void {
    const id = deps.sessionId();
    if (!id || !deps.isActive()) {
      return;
    }
    if (deps.hasUnsavedChanges()) {
      void deps.saveDraft(id).catch(() => {});
    } else if (deps.mayCommit()) {
      void deps
        .commit(id)
        .then(() => deps.onCleanClose())
        .catch(() => {});
    }
  }

  window.addEventListener("beforeunload", handleBeforeUnload);
  window.addEventListener("pagehide", handlePageHide);
  onBeforeUnmount(() => {
    window.removeEventListener("beforeunload", handleBeforeUnload);
    window.removeEventListener("pagehide", handlePageHide);
  });
}
