// The consolidated side panel (SPEC-065). The seven auxiliary panes share one
// docked, tabbed, resizable region. This composable owns the panel's state —
// dock side, width, collapse, the open tabs and the active one — and the
// open/activate/close/toggle semantics; SidePanel.vue renders the chrome and
// DocumentView renders the pane bodies into it.
//
// Each pane composable keeps its own `xPaneOpen` boolean (its data still flows
// from there). We hold those refs and keep them in two-way sync with our
// openPanes array: a composable-initiated open (inserting a comment opens the
// comments pane; a session close resets the session-only panes) flips its ref,
// which we mirror into openPanes; our own open/close flips the ref back under a
// reentrancy guard. Active-vs-open — which a single boolean cannot express —
// lives only here.

import { computed, ref, watch } from "vue";
import type { Ref } from "vue";
import type { EditorContext } from "./context";
import {
  PANE_IDS,
  clampPanelWidth,
  loadSidePanelState,
  saveSidePanelState,
  type PaneId,
  type SidePanelSide,
} from "../../lib/sidePanel";

/** A tab's identity, label and glyph (AC-30). `badge`, when present, is a live
 * count; a value greater than 0 shows a badge (AC-6). */
export interface PaneDescriptor {
  id: PaneId;
  title: string;
  icon: string;
  badge?: Readonly<Ref<number>>;
}

export interface SidePanelDeps extends EditorContext {
  /** The pane composables' own open refs, kept in two-way sync with the panel. */
  paneOpenRefs: Record<PaneId, Ref<boolean>>;
  /** Live badge counts (AC-6b): comments, track changes, accessibility. */
  badges: Partial<Record<PaneId, Readonly<Ref<number>>>>;
  /** Fired when the panel itself opens a pane — the accessibility pane rescans
   * here, since setting its ref alone does not (AC-6b). */
  onPaneOpen?: (id: PaneId) => void;
  /** Fired when the panel itself closes a pane. */
  onPaneClose?: (id: PaneId) => void;
  /** Re-evaluate fit zoom once a resize settles (SPEC-063 hook, AC-9). A
   * no-op where zoom is not wired; deliberately schedules NO pagination. */
  onResizeComplete?: () => void;
}

const PANE_TITLES: Record<PaneId, string> = {
  outline: "Outline",
  styles: "Styles",
  comments: "Comments",
  trackChanges: "Track changes",
  versionHistory: "Version history",
  accessibility: "Accessibility",
  fonts: "Fonts",
};

// Icon ids resolved by AppIcon (SPEC-058). Version history uses the shared
// `history` glyph; unknown names render nothing, so a typo degrades gracefully.
const PANE_ICONS: Record<PaneId, string> = {
  outline: "outline",
  styles: "styles",
  comments: "comments",
  trackChanges: "track-changes",
  versionHistory: "history",
  accessibility: "accessibility",
  fonts: "fonts",
};

/** A resolved tab, ready for the rail: descriptor plus the current badge count
 * (0 = no badge). */
export interface SidePanelTab {
  id: PaneId;
  title: string;
  icon: string;
  badge: number;
}

/** Typed access to SidePanel.vue's runtime `tabs` prop — the espree constraint
 * bars PropType<...> in .vue scripts, so the cast lives here (the same seam
 * asVisibleCommentThreads gives CommentsPane). */
export function asSidePanelTabs(value: unknown): SidePanelTab[] {
  return (value ?? []) as SidePanelTab[];
}

export function useSidePanel(deps: SidePanelDeps) {
  const initial = loadSidePanelState();

  // Accessibility cannot restore its scan across a reload, so a persisted-open
  // accessibility tab would show an empty pane. Force it closed on load; the
  // menu reopens it (and rescans) on demand.
  const restoredOpen: PaneId[] = initial.openPanes.filter((id) => id !== "accessibility");
  const restoredActive =
    initial.activePane !== null && restoredOpen.includes(initial.activePane)
      ? initial.activePane
      : (restoredOpen[0] ?? null);

  const side = ref<SidePanelSide>(initial.side);
  const widthPx = ref(initial.widthPx);
  const collapsed = ref(initial.collapsed);
  const openPanes = ref<PaneId[]>([...restoredOpen]);
  const activePane = ref<PaneId | null>(restoredActive);

  // Guards the panel -> composable-ref writes so our own mutations don't loop
  // back through the sync watchers below.
  let syncing = false;
  function setPaneOpenRef(id: PaneId, value: boolean): void {
    const openRef = deps.paneOpenRefs[id];
    if (!openRef || openRef.value === value) {
      return;
    }
    syncing = true;
    openRef.value = value;
    syncing = false;
  }

  // Seed the composable refs to match the restored open set before wiring the
  // sync watchers, so this initial alignment does not fire them.
  for (const id of PANE_IDS) {
    setPaneOpenRef(id, openPanes.value.includes(id));
  }

  function persist(): void {
    saveSidePanelState({
      side: side.value,
      widthPx: widthPx.value,
      collapsed: collapsed.value,
      openPanes: openPanes.value,
      activePane: activePane.value,
    });
  }
  // Write once now so the single key exists post-migration (AC-12d): a second
  // load then reads it rather than re-running the legacy seeding.
  persist();

  function openPane(id: PaneId): void {
    if (!openPanes.value.includes(id)) {
      openPanes.value = [...openPanes.value, id];
      setPaneOpenRef(id, true);
      deps.onPaneOpen?.(id);
    }
    activePane.value = id;
    // Opening a pane reveals it (AC-15a): un-collapse if the rail was collapsed.
    collapsed.value = false;
  }

  function activatePane(id: PaneId): void {
    if (openPanes.value.includes(id)) {
      activePane.value = id;
    }
  }

  function closePane(id: PaneId): void {
    const index = openPanes.value.indexOf(id);
    if (index === -1) {
      return;
    }
    const remaining = openPanes.value.filter((paneId) => paneId !== id);
    openPanes.value = remaining;
    setPaneOpenRef(id, false);
    deps.onPaneClose?.(id);
    if (activePane.value === id) {
      // The next open pane in order becomes active; the panel closes when none
      // remain (AC-24a). It never collapses while other tabs remain (AC-24b) —
      // collapse is untouched here.
      activePane.value = remaining.length > 0 ? remaining[Math.min(index, remaining.length - 1)] : null;
    }
  }

  /** Close the whole panel: remove every tab, clear the active pane (AC-7a). */
  function closePanel(): void {
    for (const id of openPanes.value) {
      setPaneOpenRef(id, false);
      deps.onPaneClose?.(id);
    }
    openPanes.value = [];
    activePane.value = null;
  }

  /** View-menu / Tools-menu toggle semantics (AC-15a/b/c). */
  function togglePane(id: PaneId): void {
    if (!openPanes.value.includes(id)) {
      openPane(id);
    } else if (activePane.value !== id) {
      activatePane(id);
    } else {
      closePane(id);
    }
  }

  // Two-way sync: a composable that flips its own open ref (insert-comment
  // opens comments; a session close resets the session-only panes) drives the
  // panel here. Our own writes are guarded out via `syncing`.
  for (const id of PANE_IDS) {
    const openRef = deps.paneOpenRefs[id];
    if (!openRef) {
      continue;
    }
    // Synchronous flush so the `syncing` guard — set and cleared around our own
    // writes in setPaneOpenRef — is still active when this fires; a default
    // (async) watcher would run after the guard cleared and re-process our own
    // mutation.
    watch(
      openRef,
      (isOpen) => {
        if (syncing) {
          return;
        }
        if (isOpen) {
          openPane(id);
        } else {
          closePane(id);
        }
      },
      { flush: "sync" },
    );
  }

  // Width is excluded here: it persists only at pointer-up (AC-8c), via
  // commitWidth / nudgeWidth.
  watch([side, collapsed, openPanes, activePane], persist);

  /** Live width during a drag — the document reflows to fill the rest (AC-8d),
   * but this does not persist. */
  function setWidth(px: number): void {
    widthPx.value = clampPanelWidth(px, window.innerWidth);
  }

  /** Pointer-up: clamp, persist the settled width (AC-8c) and re-evaluate fit
   * zoom (AC-9).
   *
   * AC-9b invariant: a panel resize schedules NO pagination. The page width is
   * fixed in centimetres and does not depend on the panel width, so this
   * handler calls nothing into the pagination engine — it only asks the zoom
   * layer (a CSS transform, SPEC-063) to re-fit, which never repaginates. */
  function commitWidth(): void {
    widthPx.value = clampPanelWidth(widthPx.value, window.innerWidth);
    persist();
    deps.onResizeComplete?.();
  }

  /** Arrow-key resize: 8px per press within bounds (AC-20a). Discrete, so it
   * commits immediately. */
  function nudgeWidth(delta: number): void {
    widthPx.value = clampPanelWidth(widthPx.value + delta, window.innerWidth);
    persist();
    deps.onResizeComplete?.();
  }

  function toggleSide(): void {
    // Moving to the left does not activate the outline (AC-3): only side flips.
    side.value = side.value === "right" ? "left" : "right";
  }

  /** Double-click the handle: collapse to the tab rail / restore the persisted
   * width (AC-10). Width is never overwritten on collapse, so expand restores
   * the last resized width. */
  function toggleCollapsed(): void {
    collapsed.value = !collapsed.value;
  }

  const tabs = computed<SidePanelTab[]>(() =>
    openPanes.value.map((id) => {
      const badgeRef = deps.badges[id];
      return {
        id,
        title: PANE_TITLES[id],
        icon: PANE_ICONS[id],
        badge: badgeRef ? badgeRef.value : 0,
      };
    }),
  );

  const activeTitle = computed(() => (activePane.value !== null ? PANE_TITLES[activePane.value] : ""));
  const panelOpen = computed(() => openPanes.value.length > 0);

  return {
    side,
    widthPx,
    collapsed,
    openPanes,
    activePane,
    tabs,
    activeTitle,
    panelOpen,
    openPane,
    activatePane,
    closePane,
    closePanel,
    togglePane,
    toggleSide,
    toggleCollapsed,
    setWidth,
    commitWidth,
    nudgeWidth,
  };
}
