import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { parseSfnt } from "./sfnt";

// Parsed against the real bundled font bytes — the same files the exporter
// fetches from /fonts at runtime — so the assertions pin actual Liberation
// metrics, not fixture approximations.
function loadFont(file: string) {
  return new Uint8Array(readFileSync(resolve(__dirname, "../../../public/fonts", file)));
}

describe("parseSfnt", () => {
  const sans = parseSfnt(loadFont("LiberationSans-Regular.ttf"));

  it("reads the head/hhea design metrics", () => {
    expect(sans.unitsPerEm).toBe(2048);
    expect(sans.ascent).toBeGreaterThan(1500);
    expect(sans.descent).toBeLessThan(0);
    const [xMin, yMin, xMax, yMax] = sans.bbox;
    expect(xMin).toBeLessThan(xMax);
    expect(yMin).toBeLessThan(yMax);
  });

  it("maps ASCII and Latin-1 code points to nonzero glyphs", () => {
    const a = sans.glyphFor(0x41); // 'A'
    const z = sans.glyphFor(0x7a); // 'z'
    const eAcute = sans.glyphFor(0xe9); // 'é'
    expect(a).toBeGreaterThan(0);
    expect(z).toBeGreaterThan(0);
    expect(eAcute).toBeGreaterThan(0);
    expect(a).not.toBe(z);
  });

  it("returns 0 (.notdef) for unmapped code points", () => {
    expect(sans.glyphFor(0xe000)).toBe(0); // private use area
    expect(sans.glyphFor(0x110000)).toBe(0); // beyond Unicode
  });

  it("reads positive advance widths that differ between glyphs", () => {
    const narrow = sans.advanceOf(sans.glyphFor(0x69)); // 'i'
    const wide = sans.advanceOf(sans.glyphFor(0x57)); // 'W'
    expect(narrow).toBeGreaterThan(0);
    expect(wide).toBeGreaterThan(narrow);
  });

  it("reads the PostScript name and cap height", () => {
    expect(sans.postScriptName).toBe("LiberationSans");
    expect(sans.capHeight).toBeGreaterThan(1000);
  });

  it("reads the italic angle of an italic face", () => {
    const italic = parseSfnt(loadFont("LiberationSans-Italic.ttf"));
    expect(italic.italicAngle).toBeLessThan(0);
    expect(sans.italicAngle).toBe(0);
  });

  it("rejects bytes that are not a TrueType font", () => {
    expect(() => parseSfnt(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toThrow(/TrueType/);
  });
});
