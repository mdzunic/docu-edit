// Theme preference persistence for the app shell's three-state appearance
// setting (SPEC-059). Pure load/save over localStorage — no Vue, no CSS var
// access — following the same tolerant, storage-guarded shape as
// lib/core/autocorrect.ts (SPEC-026) and lib/core/colors.ts (SPEC-013):
// absent/blocked/corrupt storage degrades to the session default rather than
// throwing (SPEC-003). This lives in lib/ (not lib/core/) on purpose: the
// theme system is app-shell product logic, and lib/core/** stays
// product-neutral for a future Slides host (SPEC-059 AC-68).

/** The three settable values (SPEC-059 AC-1). */
export type ThemePreference = "system" | "light" | "dark";

/** localStorage key for the theme preference (SPEC-059 AC-2). */
export const THEME_STORAGE_KEY = "docu-edit:theme";

/** localStorage key for the dim-page preference (SPEC-059 AC-72). */
export const DIM_PAGE_STORAGE_KEY = "docu-edit:dim-page";

/** Absent or unrecognised means system (SPEC-059 AC-54). */
export const DEFAULT_THEME: ThemePreference = "system";

/** Reads the stored preference, tolerating absent/blocked/unrecognised
 * storage by returning `system` (SPEC-059 AC-54, AC-56). */
export function loadThemePreference(): ThemePreference {
  try {
    const raw = localStorage.getItem(THEME_STORAGE_KEY);
    if (raw === "light" || raw === "dark" || raw === "system") {
      return raw;
    }
    return DEFAULT_THEME;
  } catch {
    return DEFAULT_THEME;
  }
}

/** Persists the preference; best-effort — a blocked write leaves the
 * in-memory setting applied (SPEC-059 AC-56, SPEC-003). */
export function saveThemePreference(pref: ThemePreference): void {
  try {
    localStorage.setItem(THEME_STORAGE_KEY, pref);
  } catch {
    // Storage unavailable (private mode / quota): the session default applies.
  }
}

/** Reads the dim-page flag: only the literal '1' means on (SPEC-059 AC-55);
 * anything else (absent, other value, blocked storage) means off. */
export function loadDimPage(): boolean {
  try {
    return localStorage.getItem(DIM_PAGE_STORAGE_KEY) === "1";
  } catch {
    return false;
  }
}

/** Persists the dim-page flag as '1' when on, clearing the key when off so
 * the stored value never lingers as anything other than '1' (SPEC-059
 * AC-55). Best-effort under the storage guard (AC-56). */
export function saveDimPage(on: boolean): void {
  try {
    if (on) {
      localStorage.setItem(DIM_PAGE_STORAGE_KEY, "1");
    } else {
      localStorage.removeItem(DIM_PAGE_STORAGE_KEY);
    }
  } catch {
    // Storage unavailable: the in-memory flag still applies this session.
  }
}
