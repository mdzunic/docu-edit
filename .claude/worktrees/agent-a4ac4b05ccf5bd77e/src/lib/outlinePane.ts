// Per-user persistence for the document outline pane's open/closed state
// (SPEC-037 AC-14, AC-15). A tiny localStorage-backed preference following the
// same pure load/save shape as lib/autocorrect.ts — no Vue, no model access.
// This is a per-USER preference, so it lives here and not on the DocumentModel
// (unlike showRuler/showFormattingMarks, which are per-document and DO get
// written to the file).

/** The localStorage key holding the pane's open/closed preference. */
export const OUTLINE_PANE_STORAGE_KEY = "docu-edit:outline-pane";

/** The pane is open by default, so a brand-new user sees it (AC-15). */
export const DEFAULT_OUTLINE_PANE_OPEN = true;

/** Load the persisted open/closed state, tolerating absent/corrupt storage
 * (private mode, quota, hand-edited JSON) by falling back to the default. */
export function loadOutlinePaneOpen(): boolean {
  try {
    const raw = localStorage.getItem(OUTLINE_PANE_STORAGE_KEY);
    if (!raw) {
      return DEFAULT_OUTLINE_PANE_OPEN;
    }
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) {
      return DEFAULT_OUTLINE_PANE_OPEN;
    }
    const open = (parsed as Record<string, unknown>).open;
    return typeof open === "boolean" ? open : DEFAULT_OUTLINE_PANE_OPEN;
  } catch {
    return DEFAULT_OUTLINE_PANE_OPEN;
  }
}

/** Persist the pane's open/closed state; best-effort (the in-memory state
 * still applies if storage is unavailable). */
export function saveOutlinePaneOpen(open: boolean): void {
  try {
    localStorage.setItem(OUTLINE_PANE_STORAGE_KEY, JSON.stringify({ open }));
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}
