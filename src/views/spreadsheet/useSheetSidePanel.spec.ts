// SPEC-129 "The side panel" — useSheetSidePanel's open/close/toggle
// semantics and its docu-edit:sheet-side-panel persistence.
import { beforeEach, describe, expect, it } from "vitest";
import { clampSheetPanelWidth, SHEET_PANES, useSheetSidePanel } from "./useSheetSidePanel";

beforeEach(() => {
  localStorage.clear();
});

describe("useSheetSidePanel (SPEC-129)", () => {
  it("starts closed with no tabs", () => {
    const panel = useSheetSidePanel();
    expect(panel.activePane.value).toBeNull();
    expect(panel.tabs.value).toEqual([]);
    expect(panel.side.value).toBe("right");
    expect(panel.collapsed.value).toBe(false);
  });

  it("open() shows the pane and adds a tab for it", () => {
    const panel = useSheetSidePanel();
    panel.open("conditionalFormats");
    expect(panel.activePane.value).toBe("conditionalFormats");
    expect(panel.tabs.value).toEqual([{ id: "conditionalFormats", ...SHEET_PANES.conditionalFormats }]);
  });

  it("close() hides the panel without dropping its tab", () => {
    const panel = useSheetSidePanel();
    panel.open("conditionalFormats");
    panel.close();
    expect(panel.activePane.value).toBeNull();
    expect(panel.tabs.value).toHaveLength(1);
  });

  it("toggle() opens a closed pane and closes the active one", () => {
    const panel = useSheetSidePanel();
    panel.toggle("conditionalFormats");
    expect(panel.activePane.value).toBe("conditionalFormats");
    panel.toggle("conditionalFormats");
    expect(panel.activePane.value).toBeNull();
  });

  it("open() also clears a collapsed panel", () => {
    const panel = useSheetSidePanel();
    panel.collapsed.value = true;
    panel.open("conditionalFormats");
    expect(panel.collapsed.value).toBe(false);
  });

  it("persists side/width/collapsed/open tabs/active pane under docu-edit:sheet-side-panel", () => {
    const panel = useSheetSidePanel();
    panel.open("conditionalFormats");
    panel.side.value = "left";
    panel.widthPx.value = 400;
    panel.collapsed.value = true;
    const stored = JSON.parse(localStorage.getItem("docu-edit:sheet-side-panel")!);
    expect(stored).toEqual({ side: "left", widthPx: 400, collapsed: true, openPanes: ["conditionalFormats"], activePane: "conditionalFormats" });
    const restored = useSheetSidePanel();
    expect(restored.side.value).toBe("left");
    expect(restored.widthPx.value).toBe(400);
    expect(restored.collapsed.value).toBe(true);
    expect(restored.activePane.value).toBe("conditionalFormats");
  });

  it("tolerates corrupt storage, falling back to the default", () => {
    localStorage.setItem("docu-edit:sheet-side-panel", "not json");
    const panel = useSheetSidePanel();
    expect(panel.activePane.value).toBeNull();
    expect(panel.side.value).toBe("right");
  });

  it("drops an unknown pane id and an incoherent active pane", () => {
    localStorage.setItem(
      "docu-edit:sheet-side-panel",
      JSON.stringify({ side: "right", widthPx: 320, collapsed: false, openPanes: ["notAPane", "conditionalFormats"], activePane: "notAPane" }),
    );
    const panel = useSheetSidePanel();
    expect(panel.tabs.value).toEqual([{ id: "conditionalFormats", ...SHEET_PANES.conditionalFormats }]);
    expect(panel.activePane.value).toBe("conditionalFormats");
  });

  it("clampSheetPanelWidth keeps the width within lib/sidePanel.ts's bounds", () => {
    expect(clampSheetPanelWidth(10)).toBeGreaterThanOrEqual(240);
    expect(clampSheetPanelWidth(100000)).toBeLessThanOrEqual(window.innerWidth / 2);
  });
});
