import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import {
  addCustomWord,
  CUSTOM_DICTIONARY_STORAGE_KEY,
  customDictionaryEntries,
  hasCustomWord,
  removeCustomWord,
  resetCustomDictionaryForTests,
  subscribeCustomDictionary,
} from "./customDictionary";

beforeEach(() => {
  localStorage.clear();
  resetCustomDictionaryForTests();
});

afterEach(() => {
  localStorage.clear();
});

describe("addCustomWord / hasCustomWord (SPEC-035 AC-7, AC-7.2)", () => {
  it("adds a word scoped to its language", () => {
    expect(addCustomWord("docuedit", "en-US")).toBe(true);
    expect(hasCustomWord("docuedit", "en-US")).toBe(true);
    expect(hasCustomWord("docuedit", "de-DE")).toBe(false);
  });

  it("is case-insensitive when checking", () => {
    addCustomWord("DocuEdit", "en-US");
    expect(hasCustomWord("docuedit", "en-US")).toBe(true);
  });

  it("rejects an empty or whitespace-only word", () => {
    expect(addCustomWord("", "en-US")).toBe(false);
    expect(addCustomWord("   ", "en-US")).toBe(false);
    expect(customDictionaryEntries()).toEqual([]);
  });

  it("rejects a duplicate for the same language", () => {
    expect(addCustomWord("docuedit", "en-US")).toBe(true);
    expect(addCustomWord("docuedit", "en-US")).toBe(false);
    expect(customDictionaryEntries()).toHaveLength(1);
  });

  it("allows the same word for a different language", () => {
    addCustomWord("chat", "fr-FR");
    expect(addCustomWord("chat", "en-US")).toBe(true);
    expect(customDictionaryEntries()).toHaveLength(2);
  });

  it("normalises the language tag's casing", () => {
    addCustomWord("docuedit", "EN-us");
    expect(hasCustomWord("docuedit", "en-US")).toBe(true);
  });
});

describe("removeCustomWord (AC-8.2)", () => {
  it("removes a matching entry", () => {
    addCustomWord("docuedit", "en-US");
    removeCustomWord("docuedit", "en-US");
    expect(hasCustomWord("docuedit", "en-US")).toBe(false);
    expect(customDictionaryEntries()).toEqual([]);
  });

  it("is a no-op for a non-matching word/language", () => {
    addCustomWord("docuedit", "en-US");
    removeCustomWord("docuedit", "de-DE");
    expect(customDictionaryEntries()).toHaveLength(1);
  });
});

describe("persistence across sessions (AC-7.2, AC-8)", () => {
  it("is readable after a fresh module load reads localStorage again", () => {
    addCustomWord("docuedit", "en-US");
    const raw = localStorage.getItem(CUSTOM_DICTIONARY_STORAGE_KEY);
    expect(raw).not.toBeNull();
    expect(JSON.parse(raw!)).toEqual([{ word: "docuedit", lang: "en-US" }]);
  });

  it("tolerates corrupt storage", () => {
    localStorage.setItem(CUSTOM_DICTIONARY_STORAGE_KEY, "{not json");
    resetCustomDictionaryForTests();
    expect(customDictionaryEntries()).toEqual([]);
  });
});

describe("subscribeCustomDictionary", () => {
  it("notifies listeners on add and remove, and unsubscribes cleanly", () => {
    const listener = vi.fn();
    const unsubscribe = subscribeCustomDictionary(listener);
    addCustomWord("docuedit", "en-US");
    expect(listener).toHaveBeenCalledTimes(1);
    removeCustomWord("docuedit", "en-US");
    expect(listener).toHaveBeenCalledTimes(2);
    unsubscribe();
    addCustomWord("again", "en-US");
    expect(listener).toHaveBeenCalledTimes(2);
  });
});
