import { describe, it, expect } from "vitest";
import { loadDictionary } from "./index";
import { checkText } from "../checker";

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

  // A dialect spelling leaking into the shared list would make both forms
  // acceptable everywhere and quietly undo AC-10. Each American form here is
  // one British English genuinely does not use — unlike "license" (the
  // British verb, beside the noun "licence") and "program" (correct in both
  // for computer programs), which are shared on purpose and would make this a
  // test of a wrong fact.
  it("keeps each US/GB pair on its own side (AC-10)", async () => {
    const us = (await loadDictionary("en-US"))!;
    const gb = (await loadDictionary("en-GB"))!;
    const pairs: Array<[string, string]> = [
      ["color", "colour"],
      ["center", "centre"],
      ["theater", "theatre"],
      ["organize", "organise"],
      ["analyze", "analyse"],
      ["defense", "defence"],
      ["catalog", "catalogue"],
      ["gray", "grey"],
      ["jewelry", "jewellery"],
      ["favorite", "favourite"],
      ["traveled", "travelled"],
    ];
    for (const [american, british] of pairs) {
      expect(us.has(american), `en-US should accept ${american}`).toBe(true);
      expect(gb.has(british), `en-GB should accept ${british}`).toBe(true);
      expect(gb.has(american), `en-GB should reject ${american}`).toBe(false);
      expect(us.has(british), `en-US should reject ${british}`).toBe(false);
    }
  });

  it("shares the spellings both dialects genuinely use", async () => {
    const us = (await loadDictionary("en-US"))!;
    const gb = (await loadDictionary("en-GB"))!;
    for (const word of ["license", "program", "check", "practice"]) {
      expect(us.has(word), `en-US should accept ${word}`).toBe(true);
      expect(gb.has(word), `en-GB should accept ${word}`).toBe(true);
    }
  });
});

// SPEC-035's whole value rests on this: a checker that underlines correct
// prose is worse than no checker, because it trains the reader to ignore
// every mark. These are ordinary sentences, not vocabulary showcases.
describe("the English dictionary against ordinary prose", () => {
  const PROSE = [
    "Following the quarterly review, the leadership team approved an updated " +
      "budget allocation for the marketing department.",
    "Please submit revised forecasts by Friday so finance can consolidate the " +
      "figures before the board presentation.",
    "I am writing to apply for the senior engineering position advertised on " +
      "your website.",
    "My background includes eight years developing distributed systems, and I " +
      "would welcome the opportunity to discuss how my experience aligns with " +
      "your goals.",
    "The quick brown fox jumps over the lazy dog.",
    "We haven't finished the client's report yet, but the read-only draft is " +
      "already shared with the team.",
  ];

  it("flags nothing in ordinary English sentences", async () => {
    const dictionary = (await loadDictionary("en-US"))!;
    const flagged = PROSE.flatMap((text) => checkText(text, dictionary).map((i) => i.word));
    expect(flagged).toEqual([]);
  });

  it("still flags actual misspellings in the same sentences", async () => {
    const dictionary = (await loadDictionary("en-US"))!;
    const issues = checkText(
      "Pleaes sumbit the revised forcast befor the bord meeting.",
      dictionary,
    );
    expect(issues.map((i) => i.word)).toEqual(["Pleaes", "sumbit", "forcast", "befor", "bord"]);
  });
});
