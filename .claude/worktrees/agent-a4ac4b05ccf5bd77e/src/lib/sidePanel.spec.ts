import { afterEach, describe, expect, it } from "vitest";
import { COMMENTS_PANE_STORAGE_KEY } from "./commentsPane";
import { OUTLINE_PANE_STORAGE_KEY } from "./outlinePane";
import { VERSION_HISTORY_PANE_STORAGE_KEY } from "./versionHistoryPane";
import {
  MIN_PANEL_WIDTH,
  SIDE_PANEL_STORAGE_KEY,
  clampPanelWidth,
  defaultSidePanelState,
  isPaneId,
  loadSidePanelState,
  saveSidePanelState,
} from "./sidePanel";

afterEach(() => {
  localStorage.clear();
});

describe("side panel persistence", () => {
  it("defaults to a closed, right-docked panel for a new user", () => {
    // No legacy keys -> nothing seeded (AC-12a): the panel starts closed so the
    // document keeps its full width until the user opens a pane.
    const state = loadSidePanelState();
    expect(state).toEqual({
      side: "right",
      widthPx: defaultSidePanelState().widthPx,
      collapsed: false,
      openPanes: [],
      activePane: null,
    });
  });

  it("round-trips the full state under the single key (AC-11)", () => {
    saveSidePanelState({
      side: "left",
      widthPx: 360,
      collapsed: true,
      openPanes: ["comments", "outline"],
      activePane: "comments",
    });
    // The state lives under exactly one key.
    expect(localStorage.getItem(SIDE_PANEL_STORAGE_KEY)).not.toBeNull();
    const state = loadSidePanelState();
    expect(state).toEqual({
      side: "left",
      widthPx: 360,
      collapsed: true,
      openPanes: ["comments", "outline"],
      activePane: "comments",
    });
  });

  it("drops unknown pane ids from openPanes and nulls an unknown activePane (AC-13)", () => {
    localStorage.setItem(
      SIDE_PANEL_STORAGE_KEY,
      JSON.stringify({
        side: "right",
        widthPx: 300,
        collapsed: false,
        openPanes: ["outline", "bogus", "comments"],
        activePane: "bogus",
      }),
    );
    const state = loadSidePanelState();
    expect(state.openPanes).toEqual(["outline", "comments"]);
    // The unknown active pane falls back to the first open tab.
    expect(state.activePane).toBe("outline");
  });

  it("falls back to defaults on corrupt storage (AC-14)", () => {
    localStorage.setItem(SIDE_PANEL_STORAGE_KEY, "not json");
    expect(loadSidePanelState()).toEqual(defaultSidePanelState());
  });
});

describe("side panel legacy migration (AC-12)", () => {
  it("seeds openPanes from truthy legacy keys in outline/comments/versionHistory order", () => {
    localStorage.setItem(VERSION_HISTORY_PANE_STORAGE_KEY, JSON.stringify({ open: true }));
    localStorage.setItem(OUTLINE_PANE_STORAGE_KEY, JSON.stringify({ open: true }));
    localStorage.setItem(COMMENTS_PANE_STORAGE_KEY, JSON.stringify({ open: true }));

    const state = loadSidePanelState();
    expect(state.openPanes).toEqual(["outline", "comments", "versionHistory"]);
    // activePane is the first seeded pane (AC-12b).
    expect(state.activePane).toBe("outline");
  });

  it("ignores legacy keys whose pane was closed", () => {
    localStorage.setItem(OUTLINE_PANE_STORAGE_KEY, JSON.stringify({ open: false }));
    localStorage.setItem(COMMENTS_PANE_STORAGE_KEY, JSON.stringify({ open: true }));

    const state = loadSidePanelState();
    expect(state.openPanes).toEqual(["comments"]);
    expect(state.activePane).toBe("comments");
  });

  it("sets activePane to null when nothing is seeded (AC-12b)", () => {
    const state = loadSidePanelState();
    expect(state.openPanes).toEqual([]);
    expect(state.activePane).toBeNull();
  });

  it("persists on migration, deletes the legacy keys and runs exactly once (AC-12c/d)", () => {
    localStorage.setItem(COMMENTS_PANE_STORAGE_KEY, JSON.stringify({ open: true }));

    const first = loadSidePanelState();
    expect(first.openPanes).toEqual(["comments"]);
    // The new key now exists and the legacy keys are gone.
    expect(localStorage.getItem(SIDE_PANEL_STORAGE_KEY)).not.toBeNull();
    expect(localStorage.getItem(COMMENTS_PANE_STORAGE_KEY)).toBeNull();

    // A second read comes from the persisted new key, not a re-migration.
    const second = loadSidePanelState();
    expect(second.openPanes).toEqual(["comments"]);
  });

  it("does not migrate once the new key exists", () => {
    saveSidePanelState({ side: "right", widthPx: 300, collapsed: false, openPanes: [], activePane: null });
    localStorage.setItem(OUTLINE_PANE_STORAGE_KEY, JSON.stringify({ open: true }));

    const state = loadSidePanelState();
    expect(state.openPanes).toEqual([]);
    // The legacy key is left untouched because migration never ran.
    expect(localStorage.getItem(OUTLINE_PANE_STORAGE_KEY)).not.toBeNull();
  });
});

describe("width clamping (AC-8a/8b)", () => {
  it("enforces the 240px minimum", () => {
    expect(clampPanelWidth(100, 2000)).toBe(MIN_PANEL_WIDTH);
  });

  it("caps at half the window width", () => {
    expect(clampPanelWidth(900, 1000)).toBe(500);
  });

  it("lets the minimum win over an even smaller half-window", () => {
    expect(clampPanelWidth(400, 300)).toBe(MIN_PANEL_WIDTH);
  });

  it("passes through a valid width unchanged", () => {
    expect(clampPanelWidth(320, 2000)).toBe(320);
  });
});

describe("isPaneId", () => {
  it("recognises the seven known ids and rejects others (AC-29)", () => {
    for (const id of ["outline", "styles", "comments", "trackChanges", "versionHistory", "accessibility", "fonts"]) {
      expect(isPaneId(id)).toBe(true);
    }
    expect(isPaneId("bogus")).toBe(false);
    expect(isPaneId(null)).toBe(false);
    expect(isPaneId(3)).toBe(false);
  });
});
