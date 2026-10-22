import { beforeEach, describe, expect, it } from "vitest";
import { nextTick } from "vue";
import { usePreferences } from "./usePreferences";
import { PREFERENCE_DEFAULTS, PREFERENCES_STORAGE_KEY } from "../lib/preferences";

// The composable is a module singleton, so these tests share its live state;
// each resets it to defaults first.
describe("usePreferences (SPEC-066)", () => {
  beforeEach(() => {
    localStorage.clear();
    usePreferences().resetPreferences();
  });

  it("exposes the current preferences reactively", () => {
    const { preferences } = usePreferences();
    expect(preferences.theme).toBe(PREFERENCE_DEFAULTS.theme);
    expect(preferences.measurementUnit).toBe(PREFERENCE_DEFAULTS.measurementUnit);
  });

  it("a setPreference persists and updates the reactive state (AC-7, AC-46)", async () => {
    const { preferences, setPreference } = usePreferences();
    setPreference("theme", "dark");
    await nextTick();
    expect(preferences.theme).toBe("dark");
    // Persisted under the single key.
    const raw = JSON.parse(localStorage.getItem(PREFERENCES_STORAGE_KEY) as string) as Record<string, unknown>;
    expect(raw.theme).toBe("dark");
  });

  it("shares one reactive object across every call site", async () => {
    const first = usePreferences();
    const second = usePreferences();
    first.setPreference("measurementUnit", "in");
    await nextTick();
    expect(second.preferences.measurementUnit).toBe("in");
  });

  it("reset restores defaults live", async () => {
    const { preferences, setPreference, resetPreferences } = usePreferences();
    setPreference("autosave", false);
    await nextTick();
    expect(preferences.autosave).toBe(false);
    resetPreferences();
    await nextTick();
    expect(preferences.autosave).toBe(PREFERENCE_DEFAULTS.autosave);
  });

  it("reflects a change made in another tab via the storage event (AC-44)", async () => {
    const { preferences } = usePreferences();
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ theme: "light" }));
    window.dispatchEvent(new StorageEvent("storage", { key: PREFERENCES_STORAGE_KEY }));
    await nextTick();
    expect(preferences.theme).toBe("light");
  });
});
