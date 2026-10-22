import { describe, it, expect, beforeEach } from "vitest";
import {
  DIM_PAGE_STORAGE_KEY,
  THEME_STORAGE_KEY,
  loadDimPage,
  loadThemePreference,
  saveDimPage,
  saveThemePreference,
} from "./themePrefs";

describe("theme preference persistence (SPEC-059)", () => {
  beforeEach(() => localStorage.clear());

  it("round-trips the three values under the documented key (AC-2/54)", () => {
    for (const value of ["system", "light", "dark"] as const) {
      saveThemePreference(value);
      expect(localStorage.getItem(THEME_STORAGE_KEY)).toBe(value);
      expect(loadThemePreference()).toBe(value);
    }
  });

  it("absent or unrecognised means system (AC-54)", () => {
    expect(loadThemePreference()).toBe("system");
    localStorage.setItem(THEME_STORAGE_KEY, "neon");
    expect(loadThemePreference()).toBe("system");
  });

  it("dim-page stores '1' when on and clears when off (AC-55/72)", () => {
    saveDimPage(true);
    expect(localStorage.getItem(DIM_PAGE_STORAGE_KEY)).toBe("1");
    expect(loadDimPage()).toBe(true);
    saveDimPage(false);
    expect(localStorage.getItem(DIM_PAGE_STORAGE_KEY)).toBeNull();
    expect(loadDimPage()).toBe(false);
  });

  it("dim-page treats anything but '1' as off (AC-55)", () => {
    localStorage.setItem(DIM_PAGE_STORAGE_KEY, "true");
    expect(loadDimPage()).toBe(false);
    localStorage.setItem(DIM_PAGE_STORAGE_KEY, "0");
    expect(loadDimPage()).toBe(false);
  });

  it("degrades to defaults when storage throws (AC-56, SPEC-003)", () => {
    const original = Object.getOwnPropertyDescriptor(window, "localStorage");
    const blocked = {
      getItem() {
        throw new Error("blocked");
      },
      setItem() {
        throw new Error("blocked");
      },
      removeItem() {
        throw new Error("blocked");
      },
    };
    Object.defineProperty(window, "localStorage", { value: blocked, configurable: true });
    try {
      expect(loadThemePreference()).toBe("system");
      expect(loadDimPage()).toBe(false);
      // Best-effort writes must not throw.
      expect(() => saveThemePreference("dark")).not.toThrow();
      expect(() => saveDimPage(true)).not.toThrow();
    } finally {
      if (original) {
        Object.defineProperty(window, "localStorage", original);
      }
    }
  });
});
