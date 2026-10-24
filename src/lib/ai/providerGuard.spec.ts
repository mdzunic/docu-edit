// SPEC-113 — the credential rules on every path that changes the stored
// provider config, not only the Labs form: a reset, another tab's edit.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { installProviderGuard } from "./providerGuard";
import { API_KEY_STORAGE_KEY, loadApiKey, saveApiKey } from "./apiKey";
import { consentGrants, grantConsent } from "./consent";
import { PREFERENCES_STORAGE_KEY, resetPreferences, resetPreferencesSession, savePreference } from "../preferences";
import { PROVIDER_CONFIG_DEFAULTS } from "./providerConfig";

const A = "https://a.example.test";
let stop = () => {};

beforeEach(() => {
  localStorage.clear();
  resetPreferencesSession();
  savePreference("modelProvider", { ...PROVIDER_CONFIG_DEFAULTS, mode: "cloud", baseUrl: `${A}/v1`, model: "m" });
  saveApiKey(A, "sk-a");
  grantConsent(A, "m", "f");
  stop = installProviderGuard();
});

afterEach(() => stop());

describe("installProviderGuard", () => {
  it("keeps the key and the grant while the config still points at them", () => {
    expect(loadApiKey(A)).toBe("sk-a");
    expect(consentGrants()).toHaveLength(1);
  });

  it("“Reset all preferences” discards the key and the grant with the origin", () => {
    resetPreferences();
    expect(localStorage.getItem(API_KEY_STORAGE_KEY)).toBeNull();
    expect(consentGrants()).toEqual([]);
  });

  it("an origin change made in another tab discards them too", () => {
    localStorage.setItem(
      PREFERENCES_STORAGE_KEY,
      JSON.stringify({ modelProvider: { mode: "cloud", baseUrl: "https://b.example.test/v1", model: "m" } }),
    );
    window.dispatchEvent(new StorageEvent("storage", { key: PREFERENCES_STORAGE_KEY }));
    expect(localStorage.getItem(API_KEY_STORAGE_KEY)).toBeNull();
    expect(consentGrants()).toEqual([]);
  });

  it("a model change clears the grant but keeps the key", () => {
    savePreference("modelProvider", { ...PROVIDER_CONFIG_DEFAULTS, mode: "cloud", baseUrl: `${A}/v1`, model: "other" });
    expect(consentGrants()).toEqual([]);
    expect(loadApiKey(A)).toBe("sk-a");
  });

  it("applies at install to a key left behind for another origin", () => {
    stop();
    saveApiKey("https://stale.example.test", "sk-stale");
    stop = installProviderGuard();
    expect(localStorage.getItem(API_KEY_STORAGE_KEY)).toBeNull();
  });
});
