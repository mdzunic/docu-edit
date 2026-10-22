// Ranking scans every word in the dictionary for every right-click, so it has
// to stay cheap against the real English list (~44k words) and not just the
// handful of words suggestions.spec.ts uses. These tests guard the two
// properties that make that possible — the length prefilter and the bounded,
// early-exiting distance function in suggestions.ts.
import { describe, it, expect } from "vitest";
import { loadDictionary } from "./dictionaries/index";
import { getSuggestions } from "./suggestions";

const TYPOS: Array<[string, string]> = [
  ["problm", "problem"],
  ["recieve", "receive"],
  ["seperate", "separate"],
  ["definately", "definitely"],
  ["occured", "occurred"],
  ["wrold", "world"],
  ["managment", "management"],
];

describe("suggestions against the full English dictionary", () => {
  it("ranks the intended correction first for common misspellings", async () => {
    const dictionary = (await loadDictionary("en-US"))!;
    for (const [typo, expected] of TYPOS) {
      expect(getSuggestions(typo, dictionary)[0], typo).toBe(expected);
    }
  });

  it("stays fast enough to answer a right-click", async () => {
    const dictionary = (await loadDictionary("en-US"))!;
    expect(dictionary.words.length).toBeGreaterThan(40000);
    for (const [typo] of TYPOS) getSuggestions(typo, dictionary); // warm

    const started = performance.now();
    for (const [typo] of TYPOS) getSuggestions(typo, dictionary);
    const perLookup = (performance.now() - started) / TYPOS.length;

    // Roughly 4 ms on a developer machine. The bound is deliberately loose —
    // it exists to catch the unbounded O(n·m) scan coming back (~70 ms at
    // this dictionary size), not to police a few milliseconds on a busy CI box.
    expect(perLookup).toBeLessThan(200);
  });
});
