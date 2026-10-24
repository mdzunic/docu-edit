// SPEC-113 — standing consent: stored per origin, revocable, scoped to one
// model and feature, and cleared when the origin or the model changes.
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
      [A]: { model: "llama", features: { "Test connection": 1000 } },
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

  it("revokes one feature's grant, or the whole origin's", () => {
    grantConsent(A, "llama", "f");
    grantConsent(A, "llama", "g");
    revokeConsent(A, "f");
    expect(consentGrants().map((grant) => grant.feature)).toEqual(["g"]);
    revokeConsent(A, "g");
    expect(localStorage.getItem(CONSENT_STORAGE_KEY)).toBeNull();
  });

  it("matches a feature by its own name only, never an inherited property", () => {
    grantConsent(A, "llama", "f");
    expect(hasConsent(A, "llama", "toString")).toBe(false);
    expect(hasConsent(A, "llama", "constructor")).toBe(false);
  });

  it("ignores corrupt stored grants", () => {
    localStorage.setItem(
      CONSENT_STORAGE_KEY,
      JSON.stringify({ [A]: { model: 3 }, [B]: "yes", "https://c.example.test": { model: "m", features: { f: "x" } } }),
    );
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

  it("is scoped to its feature: another feature is not covered, and asking for it leaves the grant alone", () => {
    grantConsent(A, "llama", "Plan a deck");
    expect(claimConsent(A, "llama", "Plan a deck")).toBe(true);
    expect(claimConsent(A, "llama", "Something else")).toBe(false);
    expect(claimConsent(A, "llama", "Plan a deck")).toBe(true);
    grantConsent(A, "llama", "Something else");
    expect(consentGrants().map((grant) => grant.feature)).toEqual(["Plan a deck", "Something else"]);
  });

  it("a model change clears every feature's grant for the origin", () => {
    grantConsent(A, "llama", "f");
    grantConsent(A, "llama", "g");
    expect(claimConsent(A, "mistral", "f")).toBe(false);
    expect(consentGrants()).toEqual([]);
    grantConsent(A, "mistral", "f");
    grantConsent(A, "llama", "g"); // granting under another model replaces the origin's grants
    expect(consentGrants().map((grant) => [grant.model, grant.feature])).toEqual([["llama", "g"]]);
  });

  it("claimConsent also drops grants for other origins and models", () => {
    grantConsent(A, "llama", "f");
    grantConsent(B, "gpt", "f");
    expect(claimConsent(B, "gpt", "f")).toBe(true);
    expect(consentGrants().map((grant) => grant.origin)).toEqual([B]);
  });
});
