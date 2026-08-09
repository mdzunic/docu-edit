// Per-user persistence for the comments margin pane's open/closed state
// (SPEC-038), the same pure load/save shape as lib/outlinePane.ts. A per-USER
// preference, not per-document — it does not live on DocumentModel.

export const COMMENTS_PANE_STORAGE_KEY = "docu-edit:comments-pane";

/** The pane is closed by default — most documents have no comments, and
 * showing an empty margin rail by default would waste space (unlike the
 * outline pane, which is useful even for a document with no headings yet). */
export const DEFAULT_COMMENTS_PANE_OPEN = false;

export function loadCommentsPaneOpen(): boolean {
  try {
    const raw = localStorage.getItem(COMMENTS_PANE_STORAGE_KEY);
    if (!raw) {
      return DEFAULT_COMMENTS_PANE_OPEN;
    }
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) {
      return DEFAULT_COMMENTS_PANE_OPEN;
    }
    const open = (parsed as Record<string, unknown>).open;
    return typeof open === "boolean" ? open : DEFAULT_COMMENTS_PANE_OPEN;
  } catch {
    return DEFAULT_COMMENTS_PANE_OPEN;
  }
}

export function saveCommentsPaneOpen(open: boolean): void {
  try {
    localStorage.setItem(COMMENTS_PANE_STORAGE_KEY, JSON.stringify({ open }));
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}
