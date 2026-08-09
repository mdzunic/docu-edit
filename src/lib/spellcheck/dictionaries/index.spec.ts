import { describe, it, expect } from "vitest";
import { loadDictionary } from "./index";

describe("loadDictionary (SPEC-035 AC-17..19)", () => {
  it("loads a bundled dictionary that recognises its own common words", async () => {
    const dictionary = await loadDictionary("en-US");
    expect(dictionary).not.toBeNull();
    expect(dictionary!.has("the")).toBe(true);
    expect(dictionary!.has("color")).toBe(true);
  });

  it("normalises the tag's casing before loading", async () => {
    const dictionary = await loadDictionary("EN-us");
    expect(dictionary!.has("the")).toBe(true);
  });

  it("distinguishes en-US and en-GB spellings (AC-10, AC-10.1)", async () => {
    const us = await loadDictionary("en-US");
    const gb = await loadDictionary("en-GB");
    expect(us!.has("color")).toBe(true);
    expect(us!.has("colour")).toBe(false);
    expect(gb!.has("colour")).toBe(true);
    expect(gb!.has("color")).toBe(false);
  });

  it("returns null synchronously for a language with no bundled dictionary (AC-19)", () => {
    expect(loadDictionary("pt-BR")).toBeNull();
  });

  it("returns the same promise on repeated calls (loaded once, cached)", () => {
    const first = loadDictionary("fr-FR");
    const second = loadDictionary("fr-FR");
    expect(first).toBe(second);
  });

  it("bundles all six required languages (AC-18)", async () => {
    for (const tag of ["en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "cs-CZ"]) {
      expect(await loadDictionary(tag)).not.toBeNull();
    }
  });
});
