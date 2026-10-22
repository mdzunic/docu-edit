import { describe, it, expect } from "vitest";
import { checkTextsInWorker } from "./spellCheckClient";

// jsdom has no global Worker, so this exercises the synchronous fallback
// path (isWorkerSupported() === false) — the worker script itself is
// covered end-to-end by e2e/SPEC-035.spec.ts, which runs in a real browser.
describe("checkTextsInWorker", () => {
  it("returns nothing for an empty batch without loading a dictionary", async () => {
    await expect(checkTextsInWorker([])).resolves.toEqual([]);
  });

  it("flags a misspelled word per item, keyed back to the caller's key", async () => {
    const results = await checkTextsInWorker([
      { key: "0", lang: "en-US", text: "This has a mispeling in it." },
      { key: "1", lang: "en-US", text: "This one is fine." },
    ]);
    const byKey = new Map(results.map((r) => [r.key, r.matches]));
    expect(byKey.get("0")?.map((m) => m.word)).toEqual(["mispeling"]);
    expect(byKey.get("1")).toEqual([]);
  });

  it("returns empty matches, not an error, for a language with no bundled dictionary", async () => {
    const results = await checkTextsInWorker([{ key: "0", lang: "xx-XX", text: "whatever text" }]);
    expect(results).toEqual([{ key: "0", matches: [] }]);
  });
});
