import { describe, it, expect } from "vitest";
import { createDictionary } from "./dictionary";
import { checkText } from "./checker";

describe("checkText (SPEC-035 AC-1)", () => {
  const dictionary = createDictionary(["the", "quick", "brown", "fox", "jumps"]);

  it("flags words the dictionary does not recognise", () => {
    const issues = checkText("the quick brwon fox", dictionary);
    expect(issues).toEqual([{ start: 10, end: 15, word: "brwon" }]);
  });

  it("is case-insensitive against the dictionary", () => {
    expect(checkText("The Quick Fox", dictionary)).toEqual([]);
  });

  it("flags nothing when every word is recognised", () => {
    expect(checkText("the quick brown fox jumps", dictionary)).toEqual([]);
  });

  it("flags every unrecognised word, ignoring the URL text itself", () => {
    const issues = checkText("visit https://example.dev/zzqxvbnm now", dictionary);
    expect(issues.map((i) => i.word)).toEqual(["visit", "now"]);
  });
});
