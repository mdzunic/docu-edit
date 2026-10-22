import { describe, it, expect } from "vitest";
import { createDictionary } from "./dictionary";
import { applyCasePattern, detectCasePattern, getSuggestions, MAX_SUGGESTIONS } from "./suggestions";

describe("getSuggestions (SPEC-035 AC-2, AC-3.1)", () => {
  const dictionary = createDictionary(["world", "word", "wood", "sword", "words", "unrelated"]);

  it("ranks a transposition typo's exact match first", () => {
    expect(getSuggestions("wrold", dictionary)[0]).toBe("world");
  });

  it("caps suggestions at five", () => {
    const big = createDictionary(["cat", "bat", "hat", "mat", "rat", "sat", "fat"]);
    expect(getSuggestions("cat", big).length).toBeLessThanOrEqual(MAX_SUGGESTIONS);
  });

  it("returns no suggestions when nothing is close (AC-3)", () => {
    expect(getSuggestions("zzqxvbnm", dictionary)).toEqual([]);
  });

  it("is case-insensitive when matching", () => {
    expect(getSuggestions("WROLD", dictionary)).toContain("world");
  });

  it("never returns duplicate suggestions", () => {
    const results = getSuggestions("wrold", dictionary);
    expect(new Set(results).size).toBe(results.length);
  });
});

describe("detectCasePattern (SPEC-035 AC-4.1)", () => {
  it("detects all-caps", () => {
    expect(detectCasePattern("WROLD")).toBe("upper");
  });

  it("detects capitalised-first-letter", () => {
    expect(detectCasePattern("Wrold")).toBe("capitalized");
  });

  it("detects lowercase", () => {
    expect(detectCasePattern("wrold")).toBe("lower");
  });

  it("detects mixed case as its own pattern", () => {
    expect(detectCasePattern("wRoLd")).toBe("mixed");
  });
});

describe("applyCasePattern (SPEC-035 AC-4.1)", () => {
  it("shapes an all-lowercase suggestion to match all-caps", () => {
    expect(applyCasePattern("world", "upper")).toBe("WORLD");
  });

  it("shapes a suggestion to capitalised-first-letter", () => {
    expect(applyCasePattern("world", "capitalized")).toBe("World");
  });

  it("lowercases for a lowercase pattern", () => {
    expect(applyCasePattern("WORLD", "lower")).toBe("world");
  });

  it("leaves a mixed-case target as the suggestion's own casing", () => {
    expect(applyCasePattern("World", "mixed")).toBe("World");
  });
});
