// Persistence, migration and validation for the consolidated side panel
// (SPEC-065). The seven auxiliary panes (outline, styles, comments, track
// changes, version history, accessibility, fonts) share ONE docked, tabbed
// region; its whole state — dock side, width, collapse, the open tabs and the
// active one — lives under a single localStorage key. On first read (the new
// key absent) it seeds from the three legacy per-pane keys (SPEC-037/038/040)
// and deletes them, so the migration runs exactly once. Pure load/save/validate
// in the SPEC-003 try/catch idiom — no Vue, no model access.

import { COMMENTS_PANE_STORAGE_KEY } from "./commentsPane";
import { OUTLINE_PANE_STORAGE_KEY } from "./outlinePane";
import { VERSION_HISTORY_PANE_STORAGE_KEY } from "./versionHistoryPane";

/** Every pane the region can host, exactly (SPEC-065 AC-29). */
export type PaneId =
  | "outline"
  | "styles"
  | "comments"
  | "trackChanges"
  | "versionHistory"
  | "accessibility"
  | "fonts";

/** The known pane ids, for validation and iteration. */
export const PANE_IDS: readonly PaneId[] = [
  "outline",
  "styles",
  "comments",
  "trackChanges",
  "versionHistory",
  "accessibility",
  "fonts",
];

export type SidePanelSide = "left" | "right";

/** The persisted shape (AC-11). */
export interface SidePanelState {
  side: SidePanelSide;
  widthPx: number;
  collapsed: boolean;
  openPanes: PaneId[];
  activePane: PaneId | null;
}

/** The single key holding the whole panel's state (AC-11). */
export const SIDE_PANEL_STORAGE_KEY = "docu-edit:side-panel";

/** Width bounds: at least 240px (AC-8a); the upper bound is half the window's
 * inner width and so is resolved at runtime (AC-8b). */
export const MIN_PANEL_WIDTH = 240;
export const DEFAULT_PANEL_WIDTH = 320;

/** The three legacy per-pane open/closed keys, in the order AC-12a seeds them
 * (outline, comments, versionHistory). */
const LEGACY_SEEDS: readonly { id: PaneId; key: string }[] = [
  { id: "outline", key: OUTLINE_PANE_STORAGE_KEY },
  { id: "comments", key: COMMENTS_PANE_STORAGE_KEY },
  { id: "versionHistory", key: VERSION_HISTORY_PANE_STORAGE_KEY },
];

export function isPaneId(value: unknown): value is PaneId {
  return typeof value === "string" && (PANE_IDS as readonly string[]).includes(value);
}

export function defaultSidePanelState(): SidePanelState {
  return { side: "right", widthPx: DEFAULT_PANEL_WIDTH, collapsed: false, openPanes: [], activePane: null };
}

/** The panel may occupy at most half the window (AC-8b). */
export function maxPanelWidth(innerWidth: number): number {
  return innerWidth / 2;
}

/** Constrain a width to [240px, innerWidth/2] (AC-8a, AC-8b). When the window
 * is so narrow that half of it is under 240px, the minimum wins. */
export function clampPanelWidth(width: number, innerWidth: number): number {
  if (!Number.isFinite(width)) {
    return MIN_PANEL_WIDTH;
  }
  const upper = Math.max(MIN_PANEL_WIDTH, maxPanelWidth(innerWidth));
  return Math.min(Math.max(width, MIN_PANEL_WIDTH), upper);
}

/** Load the persisted state, tolerating absent/blocked/corrupt storage by
 * falling back to the default (AC-14). When the new key is absent, seeds from
 * the legacy keys and deletes them (AC-12). */
export function loadSidePanelState(): SidePanelState {
  let raw: string | null = null;
  try {
    raw = localStorage.getItem(SIDE_PANEL_STORAGE_KEY);
  } catch {
    return defaultSidePanelState();
  }
  if (raw !== null) {
    try {
      return validateState(JSON.parse(raw));
    } catch {
      return defaultSidePanelState();
    }
  }
  // First read: seed from the legacy keys, then persist so the new key exists
  // and this migration runs exactly once (AC-12d) — even before the composable
  // writes any further state.
  const migrated = migrateLegacyState();
  saveSidePanelState(migrated);
  return migrated;
}

/** Persist the state; best-effort — if storage is blocked the in-memory state
 * still applies (AC-14). */
export function saveSidePanelState(state: SidePanelState): void {
  try {
    localStorage.setItem(
      SIDE_PANEL_STORAGE_KEY,
      JSON.stringify({
        side: state.side,
        widthPx: state.widthPx,
        collapsed: state.collapsed,
        openPanes: state.openPanes,
        activePane: state.activePane,
      }),
    );
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}

/** Coerce an arbitrary parsed value into a valid state. Unknown pane ids are
 * dropped from openPanes and an unknown activePane becomes null (AC-13). */
function validateState(parsed: unknown): SidePanelState {
  const base = defaultSidePanelState();
  if (typeof parsed !== "object" || parsed === null) {
    return base;
  }
  const obj = parsed as Record<string, unknown>;
  const side: SidePanelSide = obj.side === "left" ? "left" : "right";
  const widthPx =
    typeof obj.widthPx === "number" && Number.isFinite(obj.widthPx)
      ? Math.max(MIN_PANEL_WIDTH, obj.widthPx)
      : base.widthPx;
  const collapsed = obj.collapsed === true;
  const openPanes = Array.isArray(obj.openPanes)
    ? obj.openPanes.filter(isPaneId).filter((id, index, arr) => arr.indexOf(id) === index)
    : [];
  let activePane: PaneId | null = isPaneId(obj.activePane) ? obj.activePane : null;
  // An active pane that isn't open is incoherent; fall back to the first open
  // tab (or null when nothing is open).
  if (activePane === null || !openPanes.includes(activePane)) {
    activePane = openPanes[0] ?? null;
  }
  return { side, widthPx, collapsed, openPanes, activePane };
}

/** Seed openPanes from the legacy per-pane keys, then delete them so this runs
 * exactly once (AC-12). */
function migrateLegacyState(): SidePanelState {
  const state = defaultSidePanelState();
  for (const seed of LEGACY_SEEDS) {
    if (readLegacyOpen(seed.key)) {
      state.openPanes.push(seed.id);
    }
  }
  state.activePane = state.openPanes[0] ?? null;
  for (const seed of LEGACY_SEEDS) {
    try {
      localStorage.removeItem(seed.key);
    } catch {
      // Best-effort; a failed delete only risks re-reading a now-ignored key.
    }
  }
  return state;
}

/** A legacy pane counts as "open" only when its stored value is truthy
 * (AC-12a); an absent key seeds nothing, so a brand-new user starts with a
 * closed panel and a full-width document. The real legacy shape is
 * `{ open: boolean }`, so an object's `.open` decides; any other truthy scalar
 * is also treated as open. */
function readLegacyOpen(key: string): boolean {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return false;
    }
    const parsed: unknown = JSON.parse(raw);
    if (parsed !== null && typeof parsed === "object") {
      return Boolean((parsed as Record<string, unknown>).open);
    }
    return Boolean(parsed);
  } catch {
    return false;
  }
}
