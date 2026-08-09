import { afterEach, describe, expect, it } from "vitest";
import {
  DEFAULT_OUTLINE_PANE_OPEN,
  OUTLINE_PANE_STORAGE_KEY,
  loadOutlinePaneOpen,
  saveOutlinePaneOpen,
} from "./outlinePane";

afterEach(() => {
  localStorage.clear();
});

describe("outline pane persistence", () => {
  it("defaults to open for a new user (nothing stored)", () => {
    expect(DEFAULT_OUTLINE_PANE_OPEN).toBe(true);
    expect(loadOutlinePaneOpen()).toBe(true);
  });

  it("round-trips the stored state", () => {
    saveOutlinePaneOpen(false);
    expect(loadOutlinePaneOpen()).toBe(false);
    saveOutlinePaneOpen(true);
    expect(loadOutlinePaneOpen()).toBe(true);
  });

  it("falls back to the default on corrupt storage", () => {
    localStorage.setItem(OUTLINE_PANE_STORAGE_KEY, "not json");
    expect(loadOutlinePaneOpen()).toBe(true);
    localStorage.setItem(OUTLINE_PANE_STORAGE_KEY, JSON.stringify({ open: "yes" }));
    expect(loadOutlinePaneOpen()).toBe(true);
  });
});
