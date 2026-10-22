import { describe, it, expect } from "vitest";
import { createDictionary } from "./dictionary";

describe("createDictionary", () => {
  it("matches listed words case-insensitively", () => {
    const dictionary = createDictionary(["Report", "budget"]);
    expect(dictionary.has("report")).toBe(true);
    expect(dictionary.has("REPORT")).toBe(true);
    expect(dictionary.has("Budget")).toBe(true);
    expect(dictionary.has("ledger")).toBe(false);
  });

  it("matches exactly by default — no suffix guessing", () => {
    const dictionary = createDictionary(["figure", "work", "box"]);
    expect(dictionary.has("figure")).toBe(true);
    expect(dictionary.has("figures")).toBe(false);
    expect(dictionary.has("working")).toBe(false);
  });

  it("exposes the word list unchanged for suggestion ranking", () => {
    const words = ["one", "two"];
    expect(createDictionary(words).words).toBe(words);
  });
});

describe("english word shapes", () => {
  const dictionary = createDictionary(
    ["do", "is", "have", "would", "company", "employees", "client", "read", "only", "well", "known", "box"],
    { wordShapes: "english" },
  );
  const accepts = (word: string) => dictionary.has(word);

  it("accepts contractions whose stem is a listed word", () => {
    for (const word of ["don't", "isn't", "haven't", "wouldn't", "don’t"]) {
      expect(accepts(word), word).toBe(true);
    }
  });

  it("accepts singular and plural possessives", () => {
    expect(accepts("company's")).toBe(true);
    expect(accepts("client's")).toBe(true);
    expect(accepts("employees'")).toBe(true);
  });

  it("accepts a hyphenated compound only when every part is a word", () => {
    expect(accepts("read-only")).toBe(true);
    expect(accepts("well-known")).toBe(true);
    expect(accepts("read-quickly")).toBe(false);
  });

  it("does not start guessing suffixes", () => {
    // The vendored list carries real inflections; inventing them here would
    // accept "boxs" and stop catching a whole class of typo.
    expect(accepts("boxs")).toBe(false);
    expect(accepts("boxes")).toBe(false);
    expect(accepts("doing")).toBe(false);
  });

  it("rejects a contraction built on a word it does not know", () => {
    expect(accepts("shouldn't")).toBe(false);
    expect(accepts("teh's")).toBe(false);
  });

  it("never rejects a word the plain list already contained", () => {
    for (const word of ["do", "company", "box"]) {
      expect(accepts(word), word).toBe(true);
    }
  });
});
