// Per-user persistence for the document library's view mode and sort choice
// (SPEC-067). SPEC-066 (a general preferences layer) is referenced by the spec
// but not yet in the tree, so this follows the same pure localStorage load/save
// shape every other per-user preference in src/lib uses (see outlinePane.ts,
// themePrefs.ts) — no Vue, no model access. When SPEC-066 lands, this reads and
// writes migrate to it without changing callers.

/** How the library lays out its documents. */
export type LibraryViewMode = "list" | "grid";

/** What the library sorts by. "modified" is last-modified time, "size" the
 * document's byte footprint. */
export type LibrarySortKey = "modified" | "title" | "size";

export type SortDirection = "asc" | "desc";

export interface LibrarySort {
  key: LibrarySortKey;
  direction: SortDirection;
}

/** The localStorage key holding the library preferences. */
export const LIBRARY_PREFS_STORAGE_KEY = "docu-edit:library";

/** A fresh user sees the list view. */
export const DEFAULT_LIBRARY_VIEW: LibraryViewMode = "list";

/** The default sort: last modified, newest first (SPEC-067 AC-11). */
export const DEFAULT_LIBRARY_SORT: LibrarySort = {
  key: "modified",
  direction: "desc",
};

const VIEW_MODES: readonly LibraryViewMode[] = ["list", "grid"];
const SORT_KEYS: readonly LibrarySortKey[] = ["modified", "title", "size"];
const DIRECTIONS: readonly SortDirection[] = ["asc", "desc"];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

/** Load the persisted view mode, tolerating absent/corrupt storage by falling
 * back to the default. */
export function loadLibraryView(): LibraryViewMode {
  try {
    const parsed: unknown = JSON.parse(
      localStorage.getItem(LIBRARY_PREFS_STORAGE_KEY) ?? "null",
    );
    if (isRecord(parsed) && VIEW_MODES.includes(parsed.view as LibraryViewMode)) {
      return parsed.view as LibraryViewMode;
    }
  } catch {
    // Fall through to the default.
  }
  return DEFAULT_LIBRARY_VIEW;
}

/** Load the persisted sort, tolerating absent/corrupt storage. */
export function loadLibrarySort(): LibrarySort {
  try {
    const parsed: unknown = JSON.parse(
      localStorage.getItem(LIBRARY_PREFS_STORAGE_KEY) ?? "null",
    );
    if (isRecord(parsed) && isRecord(parsed.sort)) {
      const { key, direction } = parsed.sort;
      if (
        SORT_KEYS.includes(key as LibrarySortKey) &&
        DIRECTIONS.includes(direction as SortDirection)
      ) {
        return { key: key as LibrarySortKey, direction: direction as SortDirection };
      }
    }
  } catch {
    // Fall through to the default.
  }
  return { ...DEFAULT_LIBRARY_SORT };
}

// Both preferences share one storage key (one JSON object), so a write of
// either must preserve the other. Merge onto whatever is already stored.
function mergePrefs(patch: Record<string, unknown>): void {
  try {
    const current: unknown = JSON.parse(
      localStorage.getItem(LIBRARY_PREFS_STORAGE_KEY) ?? "null",
    );
    const base = isRecord(current) ? current : {};
    localStorage.setItem(
      LIBRARY_PREFS_STORAGE_KEY,
      JSON.stringify({ ...base, ...patch }),
    );
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}

/** Persist the view mode; best-effort. */
export function saveLibraryView(view: LibraryViewMode): void {
  mergePrefs({ view });
}

/** Persist the sort; best-effort. */
export function saveLibrarySort(sort: LibrarySort): void {
  mergePrefs({ sort });
}
