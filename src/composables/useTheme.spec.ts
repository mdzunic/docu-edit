import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { nextTick } from "vue";
import { resolveThemeColor, useTheme } from "./useTheme";
import { THEME_STORAGE_KEY } from "../lib/themePrefs";

// The composable is a module singleton, so these tests exercise the live
// shared state; each resets the DOM attributes and storage it touches.
describe("useTheme (SPEC-059)", () => {
  beforeEach(() => {
    localStorage.clear();
    document.documentElement.removeAttribute("data-theme");
  });
  afterEach(() => {
    // Leave the singleton in `system` so later files start clean.
    useTheme().setTheme("system");
  });

  it("sets data-theme for an explicit choice and removes it for system (AC-4/6)", async () => {
    const { setTheme } = useTheme();
    setTheme("dark");
    await nextTick();
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    setTheme("light");
    await nextTick();
    expect(document.documentElement.getAttribute("data-theme")).toBe("light");
    setTheme("system");
    await nextTick();
    expect(document.documentElement.hasAttribute("data-theme")).toBe(false);
  });

  it("persists the choice under the documented key (AC-2)", async () => {
    const { setTheme } = useTheme();
    setTheme("dark");
    await nextTick();
    expect(localStorage.getItem(THEME_STORAGE_KEY)).toBe("dark");
  });

  it("resolvedTheme collapses system to the OS value and honours explicit choices (AC-3)", async () => {
    const { setTheme, resolvedTheme } = useTheme();
    setTheme("light");
    await nextTick();
    expect(resolvedTheme.value).toBe("light");
    setTheme("dark");
    await nextTick();
    expect(resolvedTheme.value).toBe("dark");
  });

  it("toggles the dim-page flag onto <html> (AC-13)", async () => {
    const { toggleDimPage, dimPage } = useTheme();
    const before = dimPage.value;
    toggleDimPage();
    await nextTick();
    expect(dimPage.value).toBe(!before);
    expect(document.documentElement.getAttribute("data-dim-page")).toBe(dimPage.value ? "1" : "0");
    if (dimPage.value) {
      toggleDimPage();
      await nextTick();
    }
  });
});

describe("resolveThemeColor (SPEC-059 AC-61/62)", () => {
  it("returns an empty string for an unknown token", () => {
    expect(resolveThemeColor("--does-not-exist")).toBe("");
  });

  it("reads a token off documentElement, accepting names with or without --", () => {
    document.documentElement.style.setProperty("--probe-token", "#123456");
    try {
      expect(resolveThemeColor("--probe-token")).toBe("#123456");
      expect(resolveThemeColor("probe-token")).toBe("#123456");
    } finally {
      document.documentElement.style.removeProperty("--probe-token");
    }
  });
});
