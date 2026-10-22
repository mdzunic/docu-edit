import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { parseSfnt } from "../core/sfnt";
import { subsetFont } from "./subset";

function bundled(file: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(__dirname, "../../../public/fonts", file)));
}

function codepoints(text: string): number[] {
  return Array.from(text).map((ch) => ch.codePointAt(0) as number);
}

describe("subsetFont", () => {
  it("produces a smaller, still-parseable font that keeps the used glyphs", () => {
    const original = bundled("LiberationSans-Regular.ttf");
    const before = parseSfnt(original);
    const subset = subsetFont(original, codepoints("Hello, World!"));
    expect(subset.length).toBeLessThan(original.length);

    // Re-parses as a valid TrueType font.
    const after = parseSfnt(subset);
    expect(after.unitsPerEm).toBe(before.unitsPerEm);
    // The glyph ids for the used characters are unchanged (no renumbering),
    // and their advances survive.
    for (const cp of codepoints("Hello, World!")) {
      const gid = before.glyphFor(cp);
      expect(after.glyphFor(cp)).toBe(gid);
      expect(after.advanceOf(gid)).toBe(before.advanceOf(gid));
    }
  });

  it("subsets a Carlito face too (composite glyphs, accents)", () => {
    const original = bundled("Carlito-Regular.ttf");
    const subset = subsetFont(original, codepoints("Café — naïve"));
    expect(subset.length).toBeLessThan(original.length);
    const after = parseSfnt(subset);
    for (const cp of codepoints("Café")) {
      expect(after.glyphFor(cp)).toBe(parseSfnt(original).glyphFor(cp));
    }
  });

  it("returns the bytes unchanged for a non-glyf font", () => {
    const notAFont = new Uint8Array([0x4f, 0x54, 0x54, 0x4f, 0, 0, 0, 0, 0, 0, 0, 0]);
    expect(subsetFont(notAFont, codepoints("x"))).toBe(notAFont);
  });
});
