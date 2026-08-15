// Per-user persistence for the comment author identity (SPEC-038 AC "the
// author name comes from a settable identity, defaulting to 'You'"). A tiny
// localStorage-backed preference following the same pure load/save shape as
// lib/outlinePane.ts — no Vue, no model access. This is a per-USER, per-
// browser-profile preference (like the outline pane's open/closed state),
// so it deliberately does NOT live on DocumentModel: your name should stay
// the same across every document you open, not travel inside one file.

/** The localStorage key holding the author-identity preference. */
export const COMMENT_AUTHOR_STORAGE_KEY = "docu-edit:comment-author";

/** The author name shown/used until the user sets their own (AC). */
export const DEFAULT_COMMENT_AUTHOR = "You";

/** Load the persisted author name, tolerating absent/corrupt storage
 * (private mode, quota, hand-edited JSON) by falling back to the default. */
export function loadCommentAuthor(): string {
  try {
    const raw = localStorage.getItem(COMMENT_AUTHOR_STORAGE_KEY);
    if (!raw) {
      return DEFAULT_COMMENT_AUTHOR;
    }
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) {
      return DEFAULT_COMMENT_AUTHOR;
    }
    const name = (parsed as Record<string, unknown>).name;
    return typeof name === "string" && name.trim().length > 0 ? name : DEFAULT_COMMENT_AUTHOR;
  } catch {
    return DEFAULT_COMMENT_AUTHOR;
  }
}

/** Persist the author name; best-effort (the in-memory state still applies
 * if storage is unavailable). An empty/whitespace-only name resets to the
 * default rather than persisting blank comments' authorship. */
export function saveCommentAuthor(name: string): void {
  try {
    const trimmed = name.trim();
    localStorage.setItem(
      COMMENT_AUTHOR_STORAGE_KEY,
      JSON.stringify({ name: trimmed.length > 0 ? trimmed : DEFAULT_COMMENT_AUTHOR }),
    );
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}
