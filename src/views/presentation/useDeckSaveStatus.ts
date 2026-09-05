// The deck's "is my work safe" state (SPEC-010's readout, for the other
// product). The presentation editor showed nothing at all: no indicator
// existed in its DOM at any point, before, during or after an edit. The whole
// product's promise is "your documents are stored in this browser", and the
// word processor answers that continuously while the slides side left it open.
//
// Decks do not use useAutosave. That composable owns drafts, recovery, the
// caret and the scroll offset — all document-shaped, none of it true here.
// Slide edits persist write-through instead: every `commit()` writes the
// record immediately. So this tracks the writes themselves rather than a save
// cadence, which is why "All changes saved" is a claim about a write that has
// actually settled and not a decoration.
import { computed, onUnmounted, ref } from "vue";
import { isStorageAvailable, onStorageStatusChange } from "../../lib/documentStore";
import {
  SAVE_STATE_ERROR,
  SAVE_STATE_SAVED,
  SAVE_STATE_SAVING,
} from "../../composables/useAutosave";

export type DeckSaveState = "saved" | "saving" | "error";

export function useDeckSaveStatus() {
  // How many writes are in flight. A count rather than a boolean: a burst of
  // commits (drag a shape, type, nudge) overlaps, and the readout must not
  // fall back to "saved" while a later write is still going.
  const inFlight = ref(0);
  const failed = ref(!isStorageAvailable());

  const stopWatchingStorage = onStorageStatusChange(() => {
    failed.value = !isStorageAvailable();
  });
  onUnmounted(stopWatchingStorage);

  const saveState = computed<DeckSaveState>(() => {
    if (failed.value) return "error";
    return inFlight.value > 0 ? "saving" : "saved";
  });

  const saveStateText = computed(() => {
    if (failed.value) return SAVE_STATE_ERROR;
    return inFlight.value > 0 ? SAVE_STATE_SAVING : SAVE_STATE_SAVED;
  });

  /** Hand each `commit()`'s write here. Tolerates a non-promise so a stubbed
   * store (the unit specs mock `touchDocument` as a bare spy) still works. */
  function trackWrite(write: unknown): void {
    if (!(write instanceof Promise)) {
      return;
    }
    inFlight.value += 1;
    void write.finally(() => {
      inFlight.value = Math.max(0, inFlight.value - 1);
      // The write never rejects — a failure latches the store's own
      // storage-unavailable state, so read that rather than a catch here.
      failed.value = !isStorageAvailable();
    });
  }

  return { saveState, saveStateText, trackWrite };
}
