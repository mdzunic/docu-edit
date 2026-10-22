import { afterEach, describe, expect, it } from "vitest";
import {
  DEFAULT_VERSION_HISTORY_PANE_OPEN,
  VERSION_HISTORY_PANE_STORAGE_KEY,
  loadVersionHistoryPaneOpen,
  saveVersionHistoryPaneOpen,
} from "./versionHistoryPane";

afterEach(() => {
  localStorage.clear();
});

describe("version history pane persistence", () => {
  it("defaults to closed for a new user (nothing stored)", () => {
    expect(DEFAULT_VERSION_HISTORY_PANE_OPEN).toBe(false);
    expect(loadVersionHistoryPaneOpen()).toBe(false);
  });

  it("round-trips the stored state", () => {
    saveVersionHistoryPaneOpen(true);
    expect(loadVersionHistoryPaneOpen()).toBe(true);
    saveVersionHistoryPaneOpen(false);
    expect(loadVersionHistoryPaneOpen()).toBe(false);
  });

  it("falls back to the default on corrupt storage", () => {
    localStorage.setItem(VERSION_HISTORY_PANE_STORAGE_KEY, "not json");
    expect(loadVersionHistoryPaneOpen()).toBe(false);
    localStorage.setItem(VERSION_HISTORY_PANE_STORAGE_KEY, JSON.stringify({ open: "yes" }));
    expect(loadVersionHistoryPaneOpen()).toBe(false);
  });
});
