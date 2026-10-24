// SPEC-113 — standing consent: stored per origin, revocable, and invalidated
// when the origin, the model or the feature changes.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  CONSENT_STORAGE_KEY,
  claimConsent,
  consentGrants,
  grantConsent,
  hasConsent,
  onConsentChange,
  pruneConsent,
  revokeConsent,
} from "./consent";

const A = "http://localhost:11434";
const B = "https://models.example.test";

beforeEach(() => {
  localStorage.clear();
});

describe("consent storage", () => {
  it("stores a grant per origin under its own key, and nothing else", () => {
    grantConsent(A, "llama", "Test connection", 1000);
    expect(JSON.parse(localStorage.getItem(CONSENT_STORAGE_KEY) ?? "{}")).toEqual({
      [A]: { model: "llama", feature: "Test connection", grantedAt: 1000 },
    });
    expect(consentGrants()).toEqual([{ origin: A, model: "llama", feature: "Test connection", grantedAt: 1000 }]);
    expect(hasConsent(A, "llama", "Test connection")).toBe(true);
    expect(hasConsent(B, "llama", "Test connection")).toBe(false);
  });

  it("is revocable per origin, and a revoked origin asks again", () => {
    grantConsent(A, "llama", "f");
    grantConsent(B, "gpt", "f");
    revokeConsent(A);
    expect(hasConsent(A, "llama", "f")).toBe(false);
    expect(consentGrants().map((grant) => grant.origin)).toEqual([B]);
    revokeConsent(B);
    expect(localStorage.getItem(CONSENT_STORAGE_KEY)).toBeNull();
  });

  it("ignores corrupt stored grants", () => {
    localStorage.setItem(CONSENT_STORAGE_KEY, JSON.stringify({ [A]: { model: 3 }, [B]: "yes" }));
    expect(consentGrants()).toEqual([]);
    localStorage.setItem(CONSENT_STORAGE_KEY, "{not json");
    expect(consentGrants()).toEqual([]);
  });

  it("notifies listeners on every change", () => {
    const listener = vi.fn();
    const stop = onConsentChange(listener);
    grantConsent(A, "llama", "f");
    revokeConsent(A);
    expect(listener).toHaveBeenCalledTimes(2);
    stop();
    grantConsent(A, "llama", "f");
    expect(listener).toHaveBeenCalledTimes(2);
  });
});

describe("invalidation", () => {
  it("is cleared when the origin changes", () => {
    grantConsent(A, "llama", "f");
    pruneConsent(B, "llama");
    expect(consentGrants()).toEqual([]);
  });

  it("is cleared when the model changes", () => {
    grantConsent(A, "llama", "f");
    pruneConsent(A, "mistral");
    expect(consentGrants()).toEqual([]);
  });

  it("survives a prune that matches it exactly", () => {
    grantConsent(A, "llama", "f");
    pruneConsent(A, "llama");
    expect(hasConsent(A, "llama", "f")).toBe(true);
  });

  it("is cleared when another feature asks — claimConsent says no and drops the grant", () => {
    grantConsent(A, "llama", "Plan a deck");
    expect(claimConsent(A, "llama", "Plan a deck")).toBe(true);
    expect(claimConsent(A, "llama", "Something else")).toBe(false);
    expect(consentGrants()).toEqual([]);
    expect(claimConsent(A, "llama", "Plan a deck")).toBe(false);
  });

  it("claimConsent also drops grants for other origins and models", () => {
    grantConsent(A, "llama", "f");
    grantConsent(B, "gpt", "f");
    expect(claimConsent(B, "gpt", "f")).toBe(true);
    expect(consentGrants().map((grant) => grant.origin)).toEqual([B]);
  });
});
