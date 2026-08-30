import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  AUTOCORRECT_STORAGE_KEY,
} from "./core/autocorrect";
import { COMMENT_AUTHOR_STORAGE_KEY } from "./core/commentIdentity";
import { DIM_PAGE_STORAGE_KEY, THEME_STORAGE_KEY } from "./themePrefs";
import {
  PREFERENCES_STORAGE_KEY,
  PREFERENCE_DEFAULTS,
  loadPreferences,
  onPreferencesChange,
  resetPreferences,
  resetPreferencesSession,
  savePreference,
  setPreferencesStorageFailureHandler,
} from "./preferences";

function stored(): Record<string, unknown> {
  const raw = localStorage.getItem(PREFERENCES_STORAGE_KEY);
  return raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
}

describe("preferences store (SPEC-066)", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferencesSession();
  });

  describe("defaults and fallback (AC-11..16)", () => {
    it("a missing stored object falls back to every declared default (AC-11)", () => {
      expect(loadPreferences()).toEqual(PREFERENCE_DEFAULTS);
    });

    it("a truncated payload falls back per field to the default (AC-15)", () => {
      // A write cut off mid-object: JSON.parse throws, so every field defaults.
      localStorage.setItem(PREFERENCES_STORAGE_KEY, '{"theme":"dark","dimPage":tr');
      expect(loadPreferences()).toEqual(PREFERENCE_DEFAULTS);
    });

    it("a malformed payload falls back per field to the default (AC-16)", () => {
      // Valid JSON but every field the wrong type — each is rejected and
      // repaired to its default.
      localStorage.setItem(
        PREFERENCES_STORAGE_KEY,
        JSON.stringify({ theme: 42, smartQuotes: "yes", defaultFontSizePt: "big", authorName: "   " }),
      );
      expect(loadPreferences()).toEqual(PREFERENCE_DEFAULTS);
    });

    it("a non-object payload falls back to defaults (AC-12)", () => {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify("just a string"));
      expect(loadPreferences()).toEqual(PREFERENCE_DEFAULTS);
    });

    it("a partially-populated object defaults only the missing fields (AC-13)", () => {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ theme: "dark", authorName: "Ada" }));
      const prefs = loadPreferences();
      expect(prefs.theme).toBe("dark");
      expect(prefs.authorName).toBe("Ada");
      // Untouched fields keep their defaults.
      expect(prefs.measurementUnit).toBe(PREFERENCE_DEFAULTS.measurementUnit);
      expect(prefs.autosave).toBe(PREFERENCE_DEFAULTS.autosave);
    });

    it("one bad field never discards the rest of the object (AC-14)", () => {
      localStorage.setItem(
        PREFERENCES_STORAGE_KEY,
        JSON.stringify({ theme: "dark", measurementUnit: "furlongs", authorName: "Ada" }),
      );
      const prefs = loadPreferences();
      expect(prefs.theme).toBe("dark");
      expect(prefs.authorName).toBe("Ada");
      // Only the invalid field is repaired.
      expect(prefs.measurementUnit).toBe("cm");
    });

    it("rejects an out-of-range default zoom but accepts a fit mode", () => {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ defaultZoom: 9000 }));
      expect(loadPreferences().defaultZoom).toBe(PREFERENCE_DEFAULTS.defaultZoom);
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ defaultZoom: "fit-page" }));
      expect(loadPreferences().defaultZoom).toBe("fit-page");
    });
  });

  describe("single key and writes (AC-9)", () => {
    it("stores every preference under docu-edit:preferences as one JSON object", () => {
      savePreference("theme", "dark");
      savePreference("measurementUnit", "in");
      const raw = localStorage.getItem(PREFERENCES_STORAGE_KEY);
      expect(raw).not.toBeNull();
      const object = JSON.parse(raw as string) as Record<string, unknown>;
      expect(object.theme).toBe("dark");
      expect(object.measurementUnit).toBe("in");
      // Round-trips through load.
      expect(loadPreferences().theme).toBe("dark");
      expect(loadPreferences().measurementUnit).toBe("in");
    });

    it("coerces an invalid saved value to that field's default", () => {
      // @ts-expect-error deliberately invalid for the runtime guard
      savePreference("theme", "neon");
      expect(loadPreferences().theme).toBe(PREFERENCE_DEFAULTS.theme);
    });

    it("defaults the SPEC-072 snap preferences (guides on, grid off, 0.25cm)", () => {
      expect(PREFERENCE_DEFAULTS.snapToGuides).toBe(true);
      expect(PREFERENCE_DEFAULTS.snapToGrid).toBe(false);
      expect(PREFERENCE_DEFAULTS.gridSpacingCm).toBe(0.25);
    });

    it("defaults the design theme to Modernist, round-trips a choice and rejects an unknown id", () => {
      expect(PREFERENCE_DEFAULTS.palette).toBe("modernist");
      savePreference("palette", "nocturne");
      expect(loadPreferences().palette).toBe("nocturne");
      // @ts-expect-error deliberately invalid for the runtime guard
      savePreference("palette", "neon");
      expect(loadPreferences().palette).toBe(PREFERENCE_DEFAULTS.palette);
    });

    it("round-trips the snap preferences and rejects an out-of-range grid spacing", () => {
      savePreference("snapToGuides", false);
      savePreference("snapToGrid", true);
      savePreference("gridSpacingCm", 1);
      expect(loadPreferences().snapToGuides).toBe(false);
      expect(loadPreferences().snapToGrid).toBe(true);
      expect(loadPreferences().gridSpacingCm).toBe(1);
      // Zero or negative spacing would divide by zero when snapping — rejected.
      savePreference("gridSpacingCm", 0);
      expect(loadPreferences().gridSpacingCm).toBe(PREFERENCE_DEFAULTS.gridSpacingCm);
    });
  });

  describe("migration of legacy keys (AC-10)", () => {
    it("folds the four legacy keys in and removes them on first read", () => {
      localStorage.setItem(THEME_STORAGE_KEY, "dark");
      localStorage.setItem(DIM_PAGE_STORAGE_KEY, "1");
      localStorage.setItem(AUTOCORRECT_STORAGE_KEY, JSON.stringify({ smartQuotes: false, dashes: false }));
      localStorage.setItem(COMMENT_AUTHOR_STORAGE_KEY, JSON.stringify({ name: "Ada" }));

      const prefs = loadPreferences();
      expect(prefs.theme).toBe("dark");
      expect(prefs.dimPage).toBe(true);
      expect(prefs.smartQuotes).toBe(false);
      expect(prefs.smartDashes).toBe(false);
      expect(prefs.authorName).toBe("Ada");

      // The legacy keys are gone, folded into the single key.
      expect(localStorage.getItem(THEME_STORAGE_KEY)).toBeNull();
      expect(localStorage.getItem(DIM_PAGE_STORAGE_KEY)).toBeNull();
      expect(localStorage.getItem(AUTOCORRECT_STORAGE_KEY)).toBeNull();
      expect(localStorage.getItem(COMMENT_AUTHOR_STORAGE_KEY)).toBeNull();
      expect(stored().theme).toBe("dark");
    });

    it("does not migrate the excluded keys (dictionary, MRU, panes)", () => {
      localStorage.setItem("docu-edit:custom-dictionary", JSON.stringify([{ word: "docu", lang: "en-US" }]));
      localStorage.setItem("docu-edit:recent-text-colors", JSON.stringify(["#112233"]));
      localStorage.setItem("docu-edit:side-panel", JSON.stringify({ openPanes: [] }));
      loadPreferences();
      expect(localStorage.getItem("docu-edit:custom-dictionary")).not.toBeNull();
      expect(localStorage.getItem("docu-edit:recent-text-colors")).not.toBeNull();
      expect(localStorage.getItem("docu-edit:side-panel")).not.toBeNull();
    });

    it("an already-stored field wins over its legacy value", () => {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ theme: "light" }));
      localStorage.setItem(THEME_STORAGE_KEY, "dark");
      expect(loadPreferences().theme).toBe("light");
      expect(localStorage.getItem(THEME_STORAGE_KEY)).toBeNull();
    });
  });

  describe("reset (AC-28..32)", () => {
    it("restores every default and only touches the preferences key", () => {
      savePreference("theme", "dark");
      savePreference("autosave", false);
      // Foreign keys that reset must not touch.
      localStorage.setItem("docu-edit:custom-dictionary", "[]");
      localStorage.setItem("draft:doc-1", "{}");

      resetPreferences();

      expect(loadPreferences()).toEqual(PREFERENCE_DEFAULTS);
      expect(localStorage.getItem("docu-edit:custom-dictionary")).toBe("[]");
      expect(localStorage.getItem("draft:doc-1")).toBe("{}");
    });
  });

  describe("change notifications (AC-7, AC-44)", () => {
    it("fires listeners for an in-tab save and for reset", () => {
      const seen: string[] = [];
      const stop = onPreferencesChange((prefs) => seen.push(prefs.theme));
      savePreference("theme", "dark");
      resetPreferences();
      stop();
      savePreference("theme", "light");
      expect(seen).toEqual(["dark", "system"]);
    });

    it("fires for a change made in another tab via the storage event (AC-44)", () => {
      const seen: string[] = [];
      const stop = onPreferencesChange((prefs) => seen.push(prefs.measurementUnit));
      // Simulate the other tab's write, then the storage event it triggers.
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ measurementUnit: "mm" }));
      window.dispatchEvent(new StorageEvent("storage", { key: PREFERENCES_STORAGE_KEY }));
      stop();
      expect(seen).toContain("mm");
    });
  });

  describe("blocked / full storage (AC-19..21)", () => {
    let restore: PropertyDescriptor | undefined;

    afterEach(() => {
      if (restore) {
        Object.defineProperty(window, "localStorage", restore);
        restore = undefined;
      }
      resetPreferencesSession();
    });

    it("degrades to in-session preferences without throwing", () => {
      restore = Object.getOwnPropertyDescriptor(window, "localStorage");
      Object.defineProperty(window, "localStorage", {
        configurable: true,
        value: {
          getItem() {
            throw new Error("blocked");
          },
          setItem() {
            throw new Error("blocked");
          },
          removeItem() {
            throw new Error("blocked");
          },
        },
      });

      expect(() => loadPreferences()).not.toThrow();
      expect(loadPreferences()).toEqual(PREFERENCE_DEFAULTS);
      // A set this session still applies, even though it can't be persisted.
      expect(() => savePreference("theme", "dark")).not.toThrow();
      expect(loadPreferences().theme).toBe("dark");
    });

    it("shows the SPEC-003 storage warning when a write is blocked (AC-21)", () => {
      const handler = vi.fn();
      setPreferencesStorageFailureHandler(handler);
      restore = Object.getOwnPropertyDescriptor(window, "localStorage");
      Object.defineProperty(window, "localStorage", {
        configurable: true,
        value: {
          getItem: () => null,
          setItem() {
            throw new Error("full");
          },
          removeItem: () => undefined,
        },
      });

      savePreference("theme", "dark");
      expect(handler).toHaveBeenCalled();
    });
  });
});
