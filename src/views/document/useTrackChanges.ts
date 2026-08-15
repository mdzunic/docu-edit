// Track changes (SPEC-039): the recording toggle, the review pane's state
// and change list, navigation and accept/reject. Recording is session-only
// (never persisted to localStorage, unlike the comments pane's open state)
// — it always starts OFF, even for a document that already has tracked
// changes in it (AC "leaves the toggle OFF"), and resets to OFF on every
// session close (a document switch), so switching documents never leaks one
// document's recording state into another's.
import { computed, ref, watch } from "vue";
import {
  acceptAllChanges,
  acceptChange as acceptChangeCommand,
  collectChanges,
  nextChangeFrom,
  previousChangeFrom,
  rejectAllChanges,
  rejectChange as rejectChangeCommand,
  type ChangeEntry,
} from "../../lib/trackChanges";
import { loadTrackChangesAuthor, saveTrackChangesAuthor } from "../../lib/trackChangesIdentity";
import type { ModelSelection } from "../../lib/commands";
import { asTrackChangesDisplayMode } from "../../lib/documentModel";
import type { EditorContext } from "./context";

/** Typed pass-through for a `ChangeEntry[]` arriving at TrackChangesPane.vue
 * as a runtime-declared Array prop (espree bars the annotation there). */
export function asChangeEntries(value: unknown): ChangeEntry[] {
  return value as ChangeEntry[];
}

export type TrackChangesDeps = EditorContext;

export function useTrackChanges(deps: TrackChangesDeps) {
  const { doc, applyEdit, getCaret, applyCaretMove, scrollModelCaretIntoView, onSessionClose } = deps;

  // --- recording toggle (AC-1, AC-2, AC-3) ------------------------------

  const trackChangesRecording = ref(false);

  function toggleTrackChangesRecording(): void {
    trackChangesRecording.value = !trackChangesRecording.value;
  }

  /** `applyEdit`'s wrapper reads this to decide whether to run the diff
   * engine over a command's result — kept as a plain function (rather than
   * exposing the ref directly) so the call site reads as intent. */
  function isRecording(): boolean {
    return trackChangesRecording.value;
  }

  // --- author identity (SPEC-038's settable identity, AC-6) -------------

  const trackChangesAuthorName = ref(loadTrackChangesAuthor());
  watch(trackChangesAuthorName, (name) => saveTrackChangesAuthor(name));

  // --- display mode (AC-9) -----------------------------------------------

  const trackChangesDisplayMode = ref(asTrackChangesDisplayMode("all"));
  function setTrackChangesDisplayMode(mode: string): void {
    trackChangesDisplayMode.value = asTrackChangesDisplayMode(mode);
  }

  // --- review pane (AC-7) -------------------------------------------------

  const trackChangesPaneOpen = ref(false);
  function toggleTrackChangesPane(): void {
    trackChangesPaneOpen.value = !trackChangesPaneOpen.value;
  }

  const changes = computed<ChangeEntry[]>(() => (doc.value ? collectChanges(doc.value) : []));

  function currentModelSelection(): ModelSelection {
    const at = getCaret();
    return at
      ? { anchor: at, focus: at }
      : { anchor: { section: 0, block: 0, offset: 0 }, focus: { section: 0, block: 0, offset: 0 } };
  }

  /** Accept/reject one change (AC-7), routed through the ordinary
   * applyEdit/undo-history path so it is undoable (AC-10) exactly like any
   * other command. `skipTrackChanges: true` — the result IS the final state
   * a reviewer chose, not a new edit to re-diff; without it, applyEdit's own
   * diff engine (still armed while recording is on) would see the rejected
   * content disappear and immediately re-track it as a fresh deletion,
   * turning "reject" into a no-op (or worse, for a row/cell removed
   * entirely — see trackChangesDiff.ts's diffTable, which has no
   * insert-your-own-then-undo short-circuit the way run text does). */
  function acceptChange(changeId: string): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit({ model: acceptChangeCommand(model, changeId), selection }, selection, "accept change", {
      skipTrackChanges: true,
    });
  }

  function rejectChange(changeId: string): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit({ model: rejectChangeCommand(model, changeId), selection }, selection, "reject change", {
      skipTrackChanges: true,
    });
  }

  /** "Accept all" (AC-7). */
  function acceptAll(): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit({ model: acceptAllChanges(model), selection }, selection, "accept all changes", {
      skipTrackChanges: true,
    });
  }

  /** "Reject all" (AC-7). */
  function rejectAll(): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit({ model: rejectAllChanges(model), selection }, selection, "reject all changes", {
      skipTrackChanges: true,
    });
  }

  // --- navigation (AC-8) --------------------------------------------------

  async function goToChange(entry: ChangeEntry | null): Promise<void> {
    if (!entry) {
      return;
    }
    applyCaretMove(entry.at);
    await scrollModelCaretIntoView(entry.at, true);
  }

  /** Jumps to a specific change by id (review-pane row click, AC-8). */
  async function goToChangeById(changeId: string): Promise<void> {
    await goToChange(changes.value.find((entry) => entry.changeId === changeId) ?? null);
  }

  /** Next-change navigation moves the caret forward to the next tracked
   * change (AC-8), wrapping to the first change past the end of the
   * document instead of doing nothing. */
  async function goToNextChange(): Promise<void> {
    const model = doc.value;
    const caret = getCaret();
    if (!model || !caret) {
      return;
    }
    await goToChange(nextChangeFrom(model, caret) ?? changes.value[0] ?? null);
  }

  /** Previous-change navigation (AC-8), wrapping to the last change before
   * the start of the document. */
  async function goToPreviousChange(): Promise<void> {
    const model = doc.value;
    const caret = getCaret();
    if (!model || !caret) {
      return;
    }
    await goToChange(previousChangeFrom(model, caret) ?? changes.value[changes.value.length - 1] ?? null);
  }

  // --- session reset (AC-2, AC-3) -----------------------------------------
  // A newly opened document always starts with recording OFF and "All
  // markup" showing, regardless of what the previous document (or this
  // one's own persisted trackChangesRecording flag) had — see the header
  // comment.
  onSessionClose(() => {
    trackChangesRecording.value = false;
    trackChangesDisplayMode.value = asTrackChangesDisplayMode("all");
    trackChangesPaneOpen.value = false;
  });

  return {
    trackChangesRecording,
    toggleTrackChangesRecording,
    isRecording,
    trackChangesAuthorName,
    trackChangesDisplayMode,
    setTrackChangesDisplayMode,
    trackChangesPaneOpen,
    toggleTrackChangesPane,
    changes,
    acceptChange,
    rejectChange,
    acceptAll,
    rejectAll,
    goToNextChange,
    goToPreviousChange,
    goToChangeById,
  };
}

export type UseTrackChangesReturn = ReturnType<typeof useTrackChanges>;
