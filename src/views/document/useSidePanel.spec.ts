import { afterEach, describe, expect, it } from "vitest";
import { nextTick, ref } from "vue";
import type { Ref } from "vue";
import { SIDE_PANEL_STORAGE_KEY, loadSidePanelState, saveSidePanelState, type PaneId } from "../../lib/sidePanel";
import { useSidePanel, type SidePanelDeps } from "./useSidePanel";

afterEach(() => {
  localStorage.clear();
});

function makePaneRefs(): Record<PaneId, Ref<boolean>> {
  return {
    outline: ref(false),
    styles: ref(false),
    comments: ref(false),
    trackChanges: ref(false),
    versionHistory: ref(false),
    accessibility: ref(false),
    fonts: ref(false),
  };
}

function setup(extra: Partial<SidePanelDeps> = {}) {
  // Unless a test has already written a state to restore, start from an
  // explicitly empty panel (a truly fresh store would migrate the outline open).
  if (localStorage.getItem(SIDE_PANEL_STORAGE_KEY) === null) {
    saveSidePanelState({ side: "right", widthPx: 320, collapsed: false, openPanes: [], activePane: null });
  }
  const paneOpenRefs = makePaneRefs();
  const events = { opened: [] as PaneId[], closed: [] as PaneId[], resized: 0 };
  const deps = {
    paneOpenRefs,
    badges: extra.badges ?? {},
    onPaneOpen: (id: PaneId) => events.opened.push(id),
    onPaneClose: (id: PaneId) => events.closed.push(id),
    onResizeComplete: () => {
      events.resized += 1;
    },
  } as unknown as SidePanelDeps;
  const panel = useSidePanel(deps);
  return { panel, paneOpenRefs, events };
}

describe("useSidePanel open/activate/close/toggle (AC-15, AC-24)", () => {
  it("toggle opens-and-activates, then activates, then closes (AC-15a/b/c)", () => {
    const { panel, paneOpenRefs } = setup();

    panel.togglePane("comments");
    expect(panel.openPanes.value).toEqual(["comments"]);
    expect(panel.activePane.value).toBe("comments");
    expect(paneOpenRefs.comments.value).toBe(true);

    panel.togglePane("outline");
    expect(panel.openPanes.value).toEqual(["comments", "outline"]);
    expect(panel.activePane.value).toBe("outline");

    // Open but not active -> becomes active, stays open.
    panel.togglePane("comments");
    expect(panel.activePane.value).toBe("comments");
    expect(panel.openPanes.value).toEqual(["comments", "outline"]);

    // Open and active -> closes; the next pane in order becomes active (AC-24a).
    panel.togglePane("comments");
    expect(panel.openPanes.value).toEqual(["outline"]);
    expect(panel.activePane.value).toBe("outline");
    expect(paneOpenRefs.comments.value).toBe(false);
  });

  it("closing the last tab closes the panel", () => {
    const { panel } = setup();
    panel.togglePane("outline");
    panel.closePane("outline");
    expect(panel.openPanes.value).toEqual([]);
    expect(panel.activePane.value).toBeNull();
    expect(panel.panelOpen.value).toBe(false);
  });

  it("closePanel removes all tabs and clears activePane (AC-7a)", () => {
    const { panel, paneOpenRefs } = setup();
    panel.togglePane("outline");
    panel.togglePane("comments");
    panel.closePanel();
    expect(panel.openPanes.value).toEqual([]);
    expect(panel.activePane.value).toBeNull();
    expect(paneOpenRefs.outline.value).toBe(false);
    expect(paneOpenRefs.comments.value).toBe(false);
  });

  it("keeps the panel open with the previous tab active when a background tab closes (AC-24b)", () => {
    const { panel } = setup();
    panel.togglePane("outline");
    panel.togglePane("comments");
    panel.togglePane("styles"); // active = styles, order [outline, comments, styles]
    panel.closePane("comments"); // a background tab
    expect(panel.openPanes.value).toEqual(["outline", "styles"]);
    expect(panel.activePane.value).toBe("styles");
  });
});

describe("useSidePanel two-way sync with composable refs", () => {
  it("mirrors a composable-initiated open into the panel (e.g. insert comment)", () => {
    const { panel, paneOpenRefs } = setup();
    paneOpenRefs.comments.value = true;
    expect(panel.openPanes.value).toEqual(["comments"]);
    expect(panel.activePane.value).toBe("comments");
  });

  it("mirrors a composable-initiated close (e.g. session reset)", () => {
    const { panel, paneOpenRefs } = setup();
    panel.togglePane("trackChanges");
    paneOpenRefs.trackChanges.value = false;
    expect(panel.openPanes.value).toEqual([]);
    expect(panel.activePane.value).toBeNull();
  });

  it("fires onPaneOpen/onPaneClose hooks (accessibility rescan/cancel)", () => {
    const { panel, events } = setup();
    panel.togglePane("accessibility");
    expect(events.opened).toContain("accessibility");
    panel.closePane("accessibility");
    expect(events.closed).toContain("accessibility");
  });
});

describe("useSidePanel side and collapse", () => {
  it("toggles side without activating the outline (AC-3)", () => {
    const { panel } = setup();
    panel.togglePane("comments");
    expect(panel.side.value).toBe("right");
    panel.toggleSide();
    expect(panel.side.value).toBe("left");
    // Moving left must not make the outline active.
    expect(panel.activePane.value).toBe("comments");
    expect(panel.openPanes.value).not.toContain("outline");
  });

  it("toggles collapse without changing the persisted width (AC-10)", () => {
    const { panel } = setup();
    panel.togglePane("outline");
    const width = panel.widthPx.value;
    panel.toggleCollapsed();
    expect(panel.collapsed.value).toBe(true);
    expect(panel.widthPx.value).toBe(width);
    panel.toggleCollapsed();
    expect(panel.collapsed.value).toBe(false);
    expect(panel.widthPx.value).toBe(width);
  });
});

describe("useSidePanel resize (AC-8, AC-9, AC-20)", () => {
  it("live width clamps but does not persist; pointer-up persists and re-fits", async () => {
    const { panel, events } = setup();
    panel.togglePane("outline");
    await nextTick();

    panel.setWidth(50); // below the 240px minimum
    expect(panel.widthPx.value).toBe(240);
    // Not yet persisted (AC-8c).
    await nextTick();
    expect(loadSidePanelState().widthPx).not.toBe(240);

    panel.commitWidth();
    expect(events.resized).toBe(1); // fit zoom re-evaluated (AC-9)
    expect(loadSidePanelState().widthPx).toBe(240);
  });

  it("nudges by 8px within bounds and persists (AC-20a)", () => {
    const { panel, events } = setup();
    panel.togglePane("outline");
    const before = panel.widthPx.value;
    panel.nudgeWidth(8);
    expect(panel.widthPx.value).toBe(before + 8);
    expect(events.resized).toBe(1);
    expect(loadSidePanelState().widthPx).toBe(before + 8);
  });
});

describe("useSidePanel tabs and badges (AC-5, AC-6)", () => {
  it("exposes a badge count only for panes with a positive badge ref", () => {
    const commentCount = ref(3);
    const { panel } = setup({ badges: { comments: commentCount } });
    panel.togglePane("comments");
    panel.togglePane("outline");
    const commentsTab = panel.tabs.value.find((tab) => tab.id === "comments");
    const outlineTab = panel.tabs.value.find((tab) => tab.id === "outline");
    expect(commentsTab?.badge).toBe(3);
    expect(commentsTab?.title).toBe("Comments");
    expect(commentsTab?.icon).toBe("comments");
    // A pane with no badge ref reports 0 (AC-6a/6c).
    expect(outlineTab?.badge).toBe(0);
  });
});

describe("useSidePanel restore", () => {
  it("restores persisted open tabs but force-closes accessibility on load", () => {
    saveSidePanelState({
      side: "left",
      widthPx: 300,
      collapsed: false,
      openPanes: ["outline", "accessibility"],
      activePane: "accessibility",
    });
    const { panel, paneOpenRefs } = setup();
    expect(panel.openPanes.value).toEqual(["outline"]);
    expect(panel.activePane.value).toBe("outline");
    expect(paneOpenRefs.outline.value).toBe(true);
    expect(paneOpenRefs.accessibility.value).toBe(false);
    expect(panel.side.value).toBe("left");
  });
});
