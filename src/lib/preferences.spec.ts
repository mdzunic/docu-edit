import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  AUTOCORRECT_STORAGE_KEY,
} from "./core/autocorrect";
import { COMMENT_AUTHOR_STORAGE_KEY } from "./core/commentIdentity";
import { DIM_PAGE_STORAGE_KEY, THEME_STORAGE_KEY } from "./themePrefs";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALE_TAGS } from "./spreadsheet/numberFormat/locale";
import {
  PREFERENCES_STORAGE_KEY,
  PREFERENCE_DEFAULTS,
  PRESENTER_NOTES_SCALES,
  assemblySheetDefaults,
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

    it("defaults the SPEC-097 Text panel closed, round-trips it and rejects a non-boolean", () => {
      expect(PREFERENCE_DEFAULTS.presentationTextOpen).toBe(false);
      expect(loadPreferences().presentationTextOpen).toBe(false);
      savePreference("presentationTextOpen", true);
      expect(loadPreferences().presentationTextOpen).toBe(true);
      expect(stored().presentationTextOpen).toBe(true);
      // @ts-expect-error deliberately invalid for the runtime guard
      savePreference("presentationTextOpen", "yes");
      expect(loadPreferences().presentationTextOpen).toBe(false);
    });

    it("defaults the SPEC-099 slide sorter off, round-trips it and rejects a non-boolean", () => {
      expect(PREFERENCE_DEFAULTS.presentationSorterOpen).toBe(false);
      expect(loadPreferences().presentationSorterOpen).toBe(false);
      savePreference("presentationSorterOpen", true);
      expect(loadPreferences().presentationSorterOpen).toBe(true);
      expect(stored().presentationSorterOpen).toBe(true);
      // @ts-expect-error deliberately invalid for the runtime guard
      savePreference("presentationSorterOpen", 1);
      expect(loadPreferences().presentationSorterOpen).toBe(false);
    });

    it("defaults the SPEC-103 version history panel closed, round-trips it and rejects a non-boolean", () => {
      expect(PREFERENCE_DEFAULTS.presentationHistoryOpen).toBe(false);
      expect(loadPreferences().presentationHistoryOpen).toBe(false);
      savePreference("presentationHistoryOpen", true);
      expect(loadPreferences().presentationHistoryOpen).toBe(true);
      expect(stored().presentationHistoryOpen).toBe(true);
      // @ts-expect-error deliberately invalid for the runtime guard
      savePreference("presentationHistoryOpen", "open");
      expect(loadPreferences().presentationHistoryOpen).toBe(false);
    });

    it("defaults the SPEC-100 presenter notes scale to 1.25, round-trips each step and rejects anything else", () => {
      expect(PREFERENCE_DEFAULTS.presentationPresenterNotesScale).toBe(1.25);
      expect(PRESENTER_NOTES_SCALES).toEqual([1, 1.25, 1.5, 2, 2.5]);
      expect(loadPreferences().presentationPresenterNotesScale).toBe(1.25);
      for (const scale of PRESENTER_NOTES_SCALES) {
        savePreference("presentationPresenterNotesScale", scale);
        expect(loadPreferences().presentationPresenterNotesScale).toBe(scale);
        expect(stored().presentationPresenterNotesScale).toBe(scale);
      }
      // A number off the steps, a numeric string (what inSet would have taken)
      // and a non-finite value all repair to the default, field by field.
      for (const bad of [1.3, 3, 0, Number.NaN, "1.5", null]) {
        localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ presentationPresenterNotesScale: bad, theme: "dark" }));
        expect(loadPreferences().presentationPresenterNotesScale, String(bad)).toBe(1.25);
        expect(loadPreferences().theme).toBe("dark");
      }
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

describe("spreadsheetLocale (SPEC-118 \"The preference\")", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferencesSession();
  });

  it("defaults to en-US, the default spreadsheet locale", () => {
    expect(PREFERENCE_DEFAULTS.spreadsheetLocale).toBe("en-US");
    expect(PREFERENCE_DEFAULTS.spreadsheetLocale).toBe(DEFAULT_SPREADSHEET_LOCALE);
    expect(loadPreferences().spreadsheetLocale).toBe("en-US");
  });

  it("keeps any of the six tags it is stored with", () => {
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ spreadsheetLocale: tag }));
      expect(loadPreferences().spreadsheetLocale).toBe(tag);
    }
  });

  it("repairs a tag outside SPREADSHEET_LOCALE_TAGS to the default alone, every other field intact", () => {
    for (const bad of ["de-AT", "de", "DE-de", "", 407, null, ["de-DE"], { tag: "de-DE" }]) {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ spreadsheetLocale: bad, theme: "dark", authorName: "Ada" }));
      const prefs = loadPreferences();
      expect(prefs.spreadsheetLocale, JSON.stringify(bad)).toBe("en-US");
      expect(prefs.theme).toBe("dark");
      expect(prefs.authorName).toBe("Ada");
    }
  });

  it("saves a valid tag, coerces an invalid one to the default, and notifies listeners", () => {
    const seen: string[] = [];
    const off = onPreferencesChange((prefs) => seen.push(prefs.spreadsheetLocale));
    savePreference("spreadsheetLocale", "cs-CZ");
    expect(stored().spreadsheetLocale).toBe("cs-CZ");
    savePreference("spreadsheetLocale", "xx-XX" as never);
    expect(stored().spreadsheetLocale).toBe("en-US");
    off();
    expect(seen).toEqual(["cs-CZ", "en-US"]);
  });
});

describe("assembly sheet preferences (SPEC-108)", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferencesSession();
  });

  it("default to the user's page-setup defaults until chosen", () => {
    expect(PREFERENCE_DEFAULTS.assemblyPaper).toBeNull();
    expect(PREFERENCE_DEFAULTS.assemblyMargins).toBeNull();
    expect(PREFERENCE_DEFAULTS.assemblyDpi).toBe(150);
    expect(assemblySheetDefaults(loadPreferences())).toEqual({
      paper: "A4",
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
      dpi: 150,
    });
    savePreference("defaultPageSizeId", "Letter");
    savePreference("defaultMarginsId", "Narrow");
    expect(assemblySheetDefaults(loadPreferences())).toMatchObject({
      paper: "Letter",
      margins: { topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 },
    });
  });

  it("persist the user's own choices, which then win over the page-setup defaults", () => {
    savePreference("defaultPageSizeId", "Letter");
    savePreference("assemblyPaper", "Fit to content");
    savePreference("assemblyMargins", { topCm: 0, rightCm: 1, bottomCm: 2, leftCm: 3 });
    savePreference("assemblyDpi", 600);
    expect(stored()).toMatchObject({
      assemblyPaper: "Fit to content",
      assemblyMargins: { topCm: 0, rightCm: 1, bottomCm: 2, leftCm: 3 },
      assemblyDpi: 600,
    });
    resetPreferencesSession();
    expect(assemblySheetDefaults(loadPreferences())).toEqual({
      paper: "Fit to content",
      margins: { topCm: 0, rightCm: 1, bottomCm: 2, leftCm: 3 },
      dpi: 600,
    });
  });

  it("repair a corrupt value to its default, field by field", () => {
    localStorage.setItem(
      PREFERENCES_STORAGE_KEY,
      JSON.stringify({ assemblyPaper: "Custom", assemblyMargins: { topCm: -1 }, assemblyDpi: 200, theme: "dark" }),
    );
    const prefs = loadPreferences();
    expect([prefs.assemblyPaper, prefs.assemblyMargins, prefs.assemblyDpi]).toEqual([null, null, 150]);
    expect(prefs.theme).toBe("dark");
  });
});
