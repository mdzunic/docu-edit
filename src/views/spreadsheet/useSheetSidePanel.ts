// The spreadsheet view's one side panel (SPEC-129 "The side panel"). Every
// later spreadsheet pane registers here rather than mounting a panel of its
// own (SPEC-135's comments, SPEC-138's version history, SPEC-139's pivot
// field list) — this spec opens the union with its one pane,
// `conditionalFormats`; SPEC-135 adds `comments` (SheetCommentsPanel).
//
// Unlike the text editor's useSidePanel (src/views/document/useSidePanel.ts),
// this composable is NOT wired to an EditorContext or a closed pane list: the
// workbook view has no per-pane "is this feature's data open" ref to
// two-way-sync (a conditional-format pane has no independent open state of
// its own), so there is nothing useSidePanel's reentrancy guard would buy
// here. State (active pane, side, width, collapse) persists under its own
// `docu-edit:sheet-side-panel` key, reusing lib/sidePanel.ts's width bounds
// (MIN_PANEL_WIDTH, DEFAULT_PANEL_WIDTH, clampPanelWidth) — not its PaneId or
// its document-specific migration, which do not apply here.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import { clampPanelWidth, DEFAULT_PANEL_WIDTH, MIN_PANEL_WIDTH } from "../../lib/sidePanel";

/** Every pane the sheet's side panel can host; later specs widen this
 * union. */
export type SheetPaneId = "conditionalFormats" | "comments";

const SHEET_PANE_IDS: readonly SheetPaneId[] = ["conditionalFormats", "comments"];

function isSheetPaneId(value: unknown): value is SheetPaneId {
  return typeof value === "string" && (SHEET_PANE_IDS as readonly string[]).includes(value);
}

/** Each pane's tab title and AppIcon glyph. */
export const SHEET_PANES: Record<SheetPaneId, { title: string; icon: string }> = {
  conditionalFormats: { title: "Conditional Formatting", icon: "conditional-format" },
  comments: { title: "Comments", icon: "comments" },
};

export type SheetSidePanelSide = "left" | "right";

const STORAGE_KEY = "docu-edit:sheet-side-panel";

interface PersistedState {
  side: SheetSidePanelSide;
  widthPx: number;
  collapsed: boolean;
  openPanes: SheetPaneId[];
  activePane: SheetPaneId | null;
}

function defaultState(): PersistedState {
  return { side: "right", widthPx: DEFAULT_PANEL_WIDTH, collapsed: false, openPanes: [], activePane: null };
}

/** Coerces an arbitrary parsed value into a valid state (an unknown pane id
 * is dropped; an active pane that is not open falls back to the first open
 * one, or null) — the same tolerance lib/sidePanel.ts's validateState gives
 * the text editor's panel. */
function validate(parsed: unknown): PersistedState {
  const base = defaultState();
  if (typeof parsed !== "object" || parsed === null) {
    return base;
  }
  const obj = parsed as Record<string, unknown>;
  const side: SheetSidePanelSide = obj.side === "left" ? "left" : "right";
  const widthPx = typeof obj.widthPx === "number" && Number.isFinite(obj.widthPx) ? Math.max(MIN_PANEL_WIDTH, obj.widthPx) : base.widthPx;
  const collapsed = obj.collapsed === true;
  const openPanes = Array.isArray(obj.openPanes) ? obj.openPanes.filter(isSheetPaneId) : [];
  let activePane: SheetPaneId | null = isSheetPaneId(obj.activePane) ? obj.activePane : null;
  if (activePane === null || !openPanes.includes(activePane)) {
    activePane = openPanes[0] ?? null;
  }
  return { side, widthPx, collapsed, openPanes, activePane };
}

function load(): PersistedState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw === null ? defaultState() : validate(JSON.parse(raw));
  } catch {
    return defaultState();
  }
}

function save(state: PersistedState): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Persistence is best-effort; the in-memory state still applies.
  }
}

export interface SheetSidePanel {
  /** null = closed. */
  activePane: Ref<SheetPaneId | null>;
  /** Shows `id` (opening a tab for it this session if it has none yet). */
  open(id: SheetPaneId): void;
  close(): void;
  /** Opens `id`, or closes the panel when it is already the active pane. */
  toggle(id: SheetPaneId): void;
  side: Ref<SheetSidePanelSide>;
  widthPx: Ref<number>;
  collapsed: Ref<boolean>;
  /** SidePanel.vue's `tabs` prop: one entry per pane opened this session. */
  tabs: ComputedRef<{ id: SheetPaneId; title: string; icon: string; badge?: string }[]>;
}

export function useSheetSidePanel(): SheetSidePanel {
  const initial = load();
  const openPanes = ref<SheetPaneId[]>(initial.openPanes);
  const activePane = ref<SheetPaneId | null>(initial.activePane);
  const side = ref<SheetSidePanelSide>(initial.side);
  const widthPx = ref(initial.widthPx);
  const collapsed = ref(initial.collapsed);

  watch(
    [openPanes, activePane, side, widthPx, collapsed],
    () => {
      save({ side: side.value, widthPx: widthPx.value, collapsed: collapsed.value, openPanes: openPanes.value, activePane: activePane.value });
    },
    // Side, widthPx and collapsed are plain writable refs in the public API
    // (no setter to hook), so persistence can only follow their change —
    // synchronously, so a caller reading storage right after a write (tests;
    // a tab close before navigating away) never races the default microtask
    // flush.
    { flush: "sync" },
  );

  function open(id: SheetPaneId): void {
    if (!openPanes.value.includes(id)) {
      openPanes.value = [...openPanes.value, id];
    }
    activePane.value = id;
    collapsed.value = false;
  }

  function close(): void {
    activePane.value = null;
  }

  function toggle(id: SheetPaneId): void {
    if (activePane.value === id) {
      close();
    } else {
      open(id);
    }
  }

  const tabs = computed(() => openPanes.value.map((id) => ({ id, ...SHEET_PANES[id] })));

  return { activePane, open, close, toggle, side, widthPx, collapsed, tabs };
}

/** `commit-width`'s clamp, bound to the live window (SPEC-065's rule,
 * reused: at most half the window, at least MIN_PANEL_WIDTH). */
export function clampSheetPanelWidth(width: number): number {
  return clampPanelWidth(width, typeof window === "undefined" ? 1024 : window.innerWidth);
}
