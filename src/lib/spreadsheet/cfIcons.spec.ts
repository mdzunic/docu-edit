import { describe, expect, it } from "vitest";
import { CF_ICON_COLORS, CF_ICON_SETS, cfGlyphPaths, cfIconAt, ICON_SET_LABELS, type CfGlyph } from "./cfIcons";
import { ICON_SET_NAMES, iconSetSize, X14_ICON_SETS, type IconSetName } from "./conditionalFormats";

/** Every glyph a set names. */
const GLYPHS: CfGlyph[] = [
  "arrowUp", "arrowUpRight", "arrowRight", "arrowDownRight", "arrowDown", "circle", "circleRimmed", "triangleUp", "triangleDown",
  "dash", "diamond", "flag", "check", "cross", "exclaim", "checkCircle", "crossCircle", "exclaimCircle", "star", "starHalf",
  "starEmpty", "bars0", "bars1", "bars2", "bars3", "bars4", "quarter0", "quarter1", "quarter2", "quarter3", "quarter4", "box0",
  "box1", "box2", "box3", "box4", "smileHappy", "smileNeutral", "smileSad",
];

/** A set's icons as glyph/colour pairs. */
function glyphs(name: IconSetName): string[] {
  return CF_ICON_SETS[name].map((icon) => `${icon.glyph}/${icon.color}`);
}

/** How many of a glyph's paths are filled with the icon colour. */
function lit(glyph: CfGlyph): number {
  return cfGlyphPaths(glyph).filter((path) => path.paint === "icon").length;
}

/** Every number in a path, as the coordinates it draws. */
function numbers(d: string): number[] {
  return [...d.matchAll(/-?\d+(?:\.\d+)?/g)].map((match) => Number(match[0]));
}

describe("CF_ICON_SETS", () => {
  it("resolves every set name — Excel's, the x14 sets and LibreOffice's smilies", () => {
    expect(Object.keys(CF_ICON_SETS).toSorted()).toEqual([...ICON_SET_NAMES].toSorted());
    for (const name of ["3Stars", "3Triangles", "5Boxes", "NoIcons"] as IconSetName[]) {
      expect(X14_ICON_SETS.has(name)).toBe(true);
    }
    expect(CF_ICON_SETS["3Smilies"]).toHaveLength(3);
    expect(CF_ICON_SETS["3ColorSmilies"]).toHaveLength(3);
    for (const name of ICON_SET_NAMES) {
      expect(ICON_SET_LABELS[name].length).toBeGreaterThan(0);
    }
  });

  it("has one glyph per threshold of each set, and none for NoIcons", () => {
    const counts = Object.fromEntries(ICON_SET_NAMES.map((name) => [name, CF_ICON_SETS[name].length]));
    expect(counts).toEqual({
      "3Arrows": 3, "3ArrowsGray": 3, "3Flags": 3, "3TrafficLights1": 3, "3TrafficLights2": 3, "3Signs": 3, "3Symbols": 3,
      "3Symbols2": 3, "3Stars": 3, "3Triangles": 3, "3Smilies": 3, "3ColorSmilies": 3, "4Arrows": 4, "4ArrowsGray": 4,
      "4RedToBlack": 4, "4Rating": 4, "4TrafficLights": 4, "5Arrows": 5, "5ArrowsGray": 5, "5Rating": 5, "5Quarters": 5,
      "5Boxes": 5, NoIcons: 0,
    });
    for (const name of ICON_SET_NAMES) {
      if (name !== "NoIcons") {
        expect(CF_ICON_SETS[name]).toHaveLength(iconSetSize(name));
      }
    }
  });

  it("orders each set lowest → highest per the spec", () => {
    const { red, amber, green, grey, black, pink, blue, gold } = CF_ICON_COLORS;
    expect(glyphs("3Arrows")).toEqual([`arrowDown/${red}`, `arrowRight/${amber}`, `arrowUp/${green}`]);
    expect(glyphs("3ArrowsGray")).toEqual([`arrowDown/${grey}`, `arrowRight/${grey}`, `arrowUp/${grey}`]);
    expect(glyphs("3Signs")).toEqual([`diamond/${red}`, `triangleUp/${amber}`, `circle/${green}`]);
    expect(glyphs("3Symbols")).toEqual([`crossCircle/${red}`, `exclaimCircle/${amber}`, `checkCircle/${green}`]);
    expect(glyphs("3Symbols2")).toEqual([`cross/${red}`, `exclaim/${amber}`, `check/${green}`]);
    expect(glyphs("3Stars")).toEqual([`starEmpty/${gold}`, `starHalf/${gold}`, `star/${gold}`]);
    expect(glyphs("3Triangles")).toEqual([`triangleDown/${red}`, `dash/${amber}`, `triangleUp/${green}`]);
    expect(glyphs("3Smilies")).toEqual([`smileSad/${black}`, `smileNeutral/${black}`, `smileHappy/${black}`]);
    expect(glyphs("3ColorSmilies")).toEqual([`smileSad/${red}`, `smileNeutral/${amber}`, `smileHappy/${green}`]);
    expect(glyphs("4Arrows")).toEqual([`arrowDown/${red}`, `arrowDownRight/${amber}`, `arrowUpRight/${amber}`, `arrowUp/${green}`]);
    expect(glyphs("4RedToBlack")).toEqual([`circle/${black}`, `circle/${grey}`, `circle/${pink}`, `circle/${red}`]);
    expect(glyphs("4Rating")).toEqual([`bars1/${blue}`, `bars2/${blue}`, `bars3/${blue}`, `bars4/${blue}`]);
    expect(glyphs("4TrafficLights")).toEqual([`circle/${black}`, `circle/${red}`, `circle/${amber}`, `circle/${green}`]);
    expect(glyphs("5Arrows")).toEqual([`arrowDown/${red}`, `arrowDownRight/${amber}`, `arrowRight/${amber}`, `arrowUpRight/${amber}`, `arrowUp/${green}`]);
    expect(glyphs("5Rating").map((entry) => entry.split("/")[0])).toEqual(["bars0", "bars1", "bars2", "bars3", "bars4"]);
    expect(glyphs("5Quarters")).toEqual(["quarter0", "quarter1", "quarter2", "quarter3", "quarter4"].map((glyph) => `${glyph}/${black}`));
    expect(glyphs("5Boxes")).toEqual(["box0", "box1", "box2", "box3", "box4"].map((glyph) => `${glyph}/${blue}`));
  });

  it("labels each icon <colour> <shape>, distinct within its set", () => {
    expect(cfIconAt("3Arrows", 2)?.label).toBe("Green up arrow");
    expect(cfIconAt("3Stars", 1)?.label).toBe("Half star");
    expect(cfIconAt("4Rating", 2)?.label).toBe("3 of 4 bars");
    for (const name of ICON_SET_NAMES) {
      const labels = CF_ICON_SETS[name].map((icon) => icon.label);
      expect(new Set(labels).size).toBe(labels.length);
    }
    expect(cfIconAt("NoIcons", 0)).toBeUndefined();
    expect(cfIconAt("3Arrows", 3)).toBeUndefined();
  });

  it("paints only with the palette", () => {
    const palette = new Set<string>(Object.values(CF_ICON_COLORS));
    for (const name of ICON_SET_NAMES) {
      for (const icon of CF_ICON_SETS[name]) {
        expect(palette.has(icon.color)).toBe(true);
      }
    }
  });
});

describe("cfGlyphPaths", () => {
  it("draws every glyph inside the 16-unit viewBox with closed paths", () => {
    for (const glyph of GLYPHS) {
      const paths = cfGlyphPaths(glyph);
      expect(paths.length).toBeGreaterThan(0);
      for (const path of paths) {
        expect(path.d).toMatch(/^M/);
        expect(path.d).toMatch(/Z$/);
        expect(["icon", "white", "none"]).toContain(path.paint);
        for (const value of numbers(path.d)) {
          expect(value).toBeGreaterThanOrEqual(0);
          expect(value).toBeLessThanOrEqual(16);
        }
      }
    }
  });

  it("lights as many bars, quarters and boxes as the glyph says", () => {
    expect(["bars0", "bars1", "bars2", "bars3", "bars4"].map((glyph) => lit(glyph as CfGlyph))).toEqual([0, 1, 2, 3, 4]);
    expect(["box0", "box1", "box2", "box3", "box4"].map((glyph) => lit(glyph as CfGlyph))).toEqual([0, 1, 2, 3, 4]);
    expect(cfGlyphPaths("quarter0").every((path) => path.paint === "none")).toBe(true);
    expect(cfGlyphPaths("quarter4").every((path) => path.paint === "icon")).toBe(true);
    expect(cfGlyphPaths("starEmpty").every((path) => path.paint === "none")).toBe(true);
  });

  it("derives the arrows by rotation: up and down mirror each other", () => {
    const up = numbers(cfGlyphPaths("arrowUp")[0]!.d);
    const down = numbers(cfGlyphPaths("arrowDown")[0]!.d);
    expect(up).toHaveLength(down.length);
    for (let i = 0; i < up.length; i += 2) {
      expect(down[i + 1]).toBeCloseTo(16 - up[i + 1]!, 1);
    }
  });
});
