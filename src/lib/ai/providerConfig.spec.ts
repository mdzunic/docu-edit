// SPEC-113 — ModelProviderConfig in preferences: off by default with no
// endpoint and no model, repaired field by field, and never carrying a key.
import { beforeEach, describe, expect, it } from "vitest";
import { PROVIDER_CONFIG_DEFAULTS, parseProviderConfig } from "./providerConfig";
import {
  PREFERENCES_STORAGE_KEY,
  PREFERENCE_DEFAULTS,
  loadPreferences,
  resetPreferencesSession,
  savePreference,
} from "../preferences";

beforeEach(() => {
  localStorage.clear();
  resetPreferencesSession();
});

describe("ModelProviderConfig", () => {
  it("defaults to off, with no provider preconfigured", () => {
    expect(PREFERENCE_DEFAULTS.modelProvider).toEqual({
      mode: "off",
      baseUrl: "",
      model: "",
      timeoutMs: 60_000,
      maxOutputTokens: 4_096,
      temperature: 0.2,
    });
    expect(loadPreferences().modelProvider).toEqual(PROVIDER_CONFIG_DEFAULTS);
    expect(Object.isFrozen(PROVIDER_CONFIG_DEFAULTS)).toBe(true);
  });

  it("is stored in the preferences object and survives a reload", () => {
    const config = { mode: "local" as const, baseUrl: "http://localhost:1234/v1", model: "m", timeoutMs: 5_000, maxOutputTokens: 64, temperature: 1 };
    savePreference("modelProvider", config);
    resetPreferencesSession();
    expect(loadPreferences().modelProvider).toEqual(config);
    expect((JSON.parse(localStorage.getItem(PREFERENCES_STORAGE_KEY) ?? "{}") as Record<string, unknown>).modelProvider).toEqual(config);
  });

  it("repairs each bad field to its default without discarding the rest", () => {
    expect(
      parseProviderConfig({ mode: "turbo", baseUrl: 3, model: "m", timeoutMs: 10, maxOutputTokens: 1.5, temperature: 9 }),
    ).toEqual({ ...PROVIDER_CONFIG_DEFAULTS, model: "m" });
    expect(parseProviderConfig("cloud")).toBeUndefined();
    expect(parseProviderConfig(null)).toBeUndefined();
  });

  it("drops an API key someone wrote into the stored object — the key is never a preference", () => {
    localStorage.setItem(
      PREFERENCES_STORAGE_KEY,
      JSON.stringify({ modelProvider: { mode: "cloud", baseUrl: "https://x.example.test/v1", model: "m", apiKey: "sk-leak" } }),
    );
    const prefs = loadPreferences();
    expect(prefs.modelProvider).not.toHaveProperty("apiKey");
    expect(JSON.stringify(prefs)).not.toContain("sk-leak");
    savePreference("theme", "dark");
    expect(localStorage.getItem(PREFERENCES_STORAGE_KEY)).not.toContain("sk-leak");
  });
});
