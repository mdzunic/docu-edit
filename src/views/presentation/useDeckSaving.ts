// The deck's saving (SPEC-103): SPEC-010's autosave and SPEC-040's version
// history, for the presentation editor. One state machine, not a second —
// useAutosave serves decks as it serves text documents (drafts on its 2 s /
// 30 s cadence, commits at the commit points, recovery, tab close), with the
// deck's state `{ caret: null, scroll: 0, slideIndex }` — composed with the
// deck's history controller (useDeckHistory). PresentationView makes one
// `useDeckSaving({ … })` call and hooks the template up; every edit reaches
// here through useSlideRender's `onCommit` → `noteEdit`.
//
// A deck has no edit lock (core/editLock.ts stays the text editor's), so the
// session simply runs while a deck is open in this view. The play view never
// mounts this, so it never writes a draft, a commit or a version.
import { onBeforeUnmount, watch, type Ref, type ShallowRef } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { isPresentationModel } from "../../lib/documentStore";
import { useAutosave } from "../../composables/useAutosave";
import { useDeckHistory } from "./useDeckHistory";

export interface DeckSavingDeps {
  model: ShallowRef<PresentationModel | null>;
  presentationId: Ref<string>;
  currentIndex: Ref<number>;
  /** usePresentationUndo.record — a restore is one undo step. */
  record: (label?: string) => void;
  /** useSlideRender.commit. */
  commit: () => void;
  reportError: (message: string) => void;
}

function clampSlide(index: number, model: PresentationModel): number {
  return Math.min(Math.max(0, index), Math.max(0, model.slides.length - 1));
}

export function useDeckSaving(deps: DeckSavingDeps) {
  const history = useDeckHistory(deps);

  /** Set while "Discard" restores the committed deck, so its slide is
   * re-selected as a recovered draft's is; an ordinary open keeps the slide
   * the view chose (the play view's `?return` hand-back). */
  let discarding = false;

  const autosave = useAutosave({
    getCaret: () => null,
    getScroll: () => 0,
    getSlideIndex: () => deps.currentIndex.value,
    isActive: () => deps.model.value !== null,
    applyRestored: (state) => {
      if (!isPresentationModel(state.model)) {
        return;
      }
      deps.model.value = state.model;
      const reselect = (state.recovered || discarding) && state.slideIndex !== undefined;
      deps.currentIndex.value = clampSlide(reselect ? state.slideIndex! : deps.currentIndex.value, state.model);
    },
    // SPEC-040's commit-point trigger for a clean tab close.
    onCleanClose: () => history.noteCommitPoint(),
  });

  // The session follows the routed deck: it begins once that deck is loaded
  // and ends — a clean close, so it commits — when the route moves on or the
  // view unmounts. Swapping the model of the same deck (recovery, a restore)
  // is not a new session.
  let sessionId: string | null = null;
  function end(): void {
    if (sessionId !== null) {
      autosave.endSession(true);
      history.end();
      sessionId = null;
    }
  }
  watch(
    () => (deps.model.value && deps.model.value.id === deps.presentationId.value ? deps.model.value.id : null),
    (id) => {
      if (id === sessionId) {
        return;
      }
      end();
      if (id !== null) {
        sessionId = id;
        history.begin(id);
        void autosave.beginSession(id);
      }
    },
    { immediate: true },
  );
  onBeforeUnmount(end);

  /** Every finished edit (useSlideRender's `onCommit`): arms the draft
   * cadence and the version cadence. */
  function noteEdit(): void {
    autosave.markDirty();
    history.noteChange();
  }

  /** A download — SPEC-010's explicit save: commit, then SPEC-040's
   * commit-point snapshot. */
  function commitPoint(): void {
    void autosave.commitNow().then(() => history.noteCommitPoint());
  }

  /** "Discard": the committed deck back, its slide re-selected. */
  function discardRecovered(): void {
    discarding = true;
    try {
      autosave.discardRecovered();
    } finally {
      discarding = false;
    }
  }

  return {
    saveState: autosave.displaySaveState,
    saveStateText: autosave.saveStateText,
    recovery: autosave.recovery,
    recoveryLabel: autosave.recoveryLabel,
    keepRecovered: autosave.keepRecovered,
    discardRecovered,
    noteEdit,
    commitPoint,
    flushForClose: autosave.flushForClose,
    history,
  };
}

export type DeckSaving = ReturnType<typeof useDeckSaving>;
