import { describe, it, expect } from "vitest";
import { tokenizeForSpellCheck } from "./tokenize";

describe("tokenizeForSpellCheck (SPEC-035)", () => {
  it("splits plain text into word tokens with correct offsets", () => {
    const tokens = tokenizeForSpellCheck("The quick fox");
    expect(tokens.map((t) => t.text)).toEqual(["The", "quick", "fox"]);
    expect(tokens[1]).toEqual({ text: "quick", start: 4, end: 9 });
  });

  it("keeps a contraction as one token", () => {
    expect(tokenizeForSpellCheck("don't stop").map((t) => t.text)).toEqual(["don't", "stop"]);
  });

  it("keeps a hyphenated compound as one token", () => {
    expect(tokenizeForSpellCheck("well-known fact").map((t) => t.text)).toEqual([
      "well-known",
      "fact",
    ]);
  });

  it("excludes a URL entirely, including its path segments (AC-12)", () => {
    const tokens = tokenizeForSpellCheck("visit https://example.dev/well-known now");
    expect(tokens.map((t) => t.text)).toEqual(["visit", "now"]);
  });

  it("excludes a www.-prefixed URL with no scheme", () => {
    const tokens = tokenizeForSpellCheck("see www.example.com for more");
    expect(tokens.map((t) => t.text)).toEqual(["see", "for", "more"]);
  });

  it("excludes an email address (AC-12)", () => {
    const tokens = tokenizeForSpellCheck("contact jane.doe@example.com today");
    expect(tokens.map((t) => t.text)).toEqual(["contact", "today"]);
  });

  it("skips bare numbers and punctuation", () => {
    expect(tokenizeForSpellCheck("It costs $42.50, really!").map((t) => t.text)).toEqual([
      "It",
      "costs",
      "really",
    ]);
  });

  it("returns no tokens for an empty or whitespace-only string", () => {
    expect(tokenizeForSpellCheck("")).toEqual([]);
    expect(tokenizeForSpellCheck("   \t")).toEqual([]);
  });

  it("handles accented and non-Latin letters as word characters", () => {
    expect(tokenizeForSpellCheck("café naïve").map((t) => t.text)).toEqual(["café", "naïve"]);
  });
});
