import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { nextTick } from "vue";
import { resolveThemeColor, useTheme } from "./useTheme";
import { PREFERENCES_STORAGE_KEY } from "../lib/preferences";

// The composable is a module singleton, so these tests exercise the live
// shared state; each resets the DOM attributes and storage it touches.
describe("useTheme (SPEC-059)", () => {
  beforeEach(() => {
    localStorage.clear();
    document.documentElement.removeAttribute("data-theme");
    document.documentElement.removeAttribute("data-palette");
  });
  afterEach(() => {
    // Leave the singleton in `system` / the default palette so later files
    // start clean.
    useTheme().setTheme("system");
    useTheme().setPalette("modernist");
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

  it("persists the choice in the preferences store (AC-2; SPEC-066 folds theme in)", async () => {
    const { setTheme } = useTheme();
    setTheme("dark");
    await nextTick();
    const prefs = JSON.parse(localStorage.getItem(PREFERENCES_STORAGE_KEY) as string) as { theme?: string };
    expect(prefs.theme).toBe("dark");
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

  it("sets data-palette for a non-default design theme and removes it for the default", async () => {
    const { setPalette } = useTheme();
    setPalette("nocturne");
    await nextTick();
    expect(document.documentElement.getAttribute("data-palette")).toBe("nocturne");
    setPalette("petrol");
    await nextTick();
    expect(document.documentElement.getAttribute("data-palette")).toBe("petrol");
    setPalette("modernist");
    await nextTick();
    expect(document.documentElement.hasAttribute("data-palette")).toBe(false);
  });

  it("persists the palette in the preferences store", async () => {
    const { setPalette } = useTheme();
    setPalette("pine");
    await nextTick();
    const prefs = JSON.parse(localStorage.getItem(PREFERENCES_STORAGE_KEY) as string) as {
      palette?: string;
    };
    expect(prefs.palette).toBe("pine");
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
