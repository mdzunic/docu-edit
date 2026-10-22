// What the app bar says about whether your work is safe (SPEC-010), and the
// recovery bar's notice beside it. Product-neutral: the word processor's
// autosave and the deck's write-through commits both report through these, so
// the two halves of the app cannot describe the same condition differently.
//
// Extracted from useAutosave, which is over its size budget and whose rule is
// to extract rather than raise it. Copy this visible is worth a test.
//
// The recovery bar's own sentence lives in lib/format.ts instead: it needs
// formatRelativeTime, and core may import only core siblings.

/** The single failure reason SPEC-010 defines: quota and write failures both
 * surface as a full store, with a download link as the escape hatch. */
export const STORAGE_FULL_REASON = "browser storage is full";

export const SAVE_STATE_SAVING = "Saving…";
export const SAVE_STATE_SAVED = "All changes saved";
export const SAVE_STATE_ERROR = `Not saved — ${STORAGE_FULL_REASON}`;

/** Shown while a recovered draft is waiting for Keep or Discard.
 *
 * The bar used to read "All changes saved" throughout that window, directly
 * beside a banner announcing recovered UNSAVED changes — two statements about
 * the same document, on screen together, contradicting each other. And the
 * banner was the truthful one: `commitNow` and `endSession` both stand down
 * while recovery is open, so nothing on screen is being written. */
export const SAVE_STATE_RECOVERED = "Not saved yet";

export type SaveState = "saved" | "saving" | "error" | "recovered";

/** The words for a state. */
export function saveStateText(state: SaveState): string {
  switch (state) {
    case "saving":
      return SAVE_STATE_SAVING;
    case "error":
      return SAVE_STATE_ERROR;
    case "recovered":
      return SAVE_STATE_RECOVERED;
    default:
      return SAVE_STATE_SAVED;
  }
}
