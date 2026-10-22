// Per-user persistence for the version history pane's open/closed state
// (SPEC-040), the same pure load/save shape as lib/commentsPane.ts. A
// per-USER preference, not per-document — it does not live on DocumentModel.

export const VERSION_HISTORY_PANE_STORAGE_KEY = "docu-edit:version-history-pane";

export const DEFAULT_VERSION_HISTORY_PANE_OPEN = false;

export function loadVersionHistoryPaneOpen(): boolean {
  try {
    const raw = localStorage.getItem(VERSION_HISTORY_PANE_STORAGE_KEY);
    if (!raw) {
      return DEFAULT_VERSION_HISTORY_PANE_OPEN;
    }
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) {
      return DEFAULT_VERSION_HISTORY_PANE_OPEN;
    }
    const open = (parsed as Record<string, unknown>).open;
    return typeof open === "boolean" ? open : DEFAULT_VERSION_HISTORY_PANE_OPEN;
  } catch {
    return DEFAULT_VERSION_HISTORY_PANE_OPEN;
  }
}

export function saveVersionHistoryPaneOpen(open: boolean): void {
  try {
    localStorage.setItem(VERSION_HISTORY_PANE_STORAGE_KEY, JSON.stringify({ open }));
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}
