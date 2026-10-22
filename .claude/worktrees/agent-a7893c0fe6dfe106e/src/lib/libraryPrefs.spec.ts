import { describe, it, expect, beforeEach } from "vitest";
import {
  DEFAULT_LIBRARY_SORT,
  DEFAULT_LIBRARY_VIEW,
  LIBRARY_PREFS_STORAGE_KEY,
  loadLibrarySort,
  loadLibraryView,
  saveLibrarySort,
  saveLibraryView,
} from "./libraryPrefs";

beforeEach(() => {
  localStorage.clear();
});

describe("libraryPrefs (SPEC-067 AC-2, AC-15)", () => {
  it("returns defaults when nothing is stored", () => {
    expect(loadLibraryView()).toBe(DEFAULT_LIBRARY_VIEW);
    expect(loadLibrarySort()).toEqual(DEFAULT_LIBRARY_SORT);
  });

  it("round-trips the view mode", () => {
    saveLibraryView("grid");
    expect(loadLibraryView()).toBe("grid");
  });

  it("round-trips the sort", () => {
    saveLibrarySort({ key: "title", direction: "asc" });
    expect(loadLibrarySort()).toEqual({ key: "title", direction: "asc" });
  });

  it("persists view and sort independently under one key", () => {
    saveLibraryView("grid");
    saveLibrarySort({ key: "size", direction: "desc" });
    expect(loadLibraryView()).toBe("grid");
    expect(loadLibrarySort()).toEqual({ key: "size", direction: "desc" });
  });

  it("falls back to defaults on corrupt values", () => {
    localStorage.setItem(
      LIBRARY_PREFS_STORAGE_KEY,
      JSON.stringify({ view: "nonsense", sort: { key: "bogus", direction: "x" } }),
    );
    expect(loadLibraryView()).toBe(DEFAULT_LIBRARY_VIEW);
    expect(loadLibrarySort()).toEqual(DEFAULT_LIBRARY_SORT);
  });
});
