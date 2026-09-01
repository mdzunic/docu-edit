// SPEC-094 AC: the shared theme model — the twelve slots, the built-in
// catalogue pinned value-by-value, `resolveColor` per transform against known
// PowerPoint values, `resolveFont`, and each built-in's dark/light contrast.
import { describe, expect, it } from "vitest";
import {
  BUILT_IN_THEMES,
  DEFAULT_THEME,
  hexToRgb,
  isThemeColor,
  isThemeFont,
  resolveColor,
  resolveFont,
  THEME_COLOR_SLOTS,
  THEME_SWATCH_SLOTS,
  THEME_SWATCH_VARIANTS,
  type DocumentTheme,
} from "./theme";

const SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

describe("the theme model", () => {
  it("declares the twelve slots in scheme order", () => {
    expect(THEME_COLOR_SLOTS).toEqual(SLOTS);
  });

  it("pins DEFAULT_THEME to the Office scheme", () => {
    expect(DEFAULT_THEME.name).toBe("Office");
    expect(DEFAULT_THEME.colors).toEqual({
      dk1: "#000000",
      lt1: "#ffffff",
      dk2: "#44546a",
      lt2: "#e7e6e6",
      accent1: "#4472c4",
      accent2: "#ed7d31",
      accent3: "#a5a5a5",
      accent4: "#ffc000",
      accent5: "#5b9bd5",
      accent6: "#70ad47",
      hlink: "#0563c1",
      folHlink: "#954f72",
    });
    expect(DEFAULT_THEME.fonts).toEqual({ major: "Calibri Light", minor: "Calibri" });
  });

  it("offers six built-ins, Office first, each complete", () => {
    expect(BUILT_IN_THEMES).toHaveLength(6);
    expect(BUILT_IN_THEMES[0]).toBe(DEFAULT_THEME);
    expect(BUILT_IN_THEMES.map((theme) => theme.name)).toEqual([
      "Office",
      "Facet",
      "Ion",
      "Retrospect",
      "Slice",
      "Wisp",
    ]);
    for (const theme of BUILT_IN_THEMES) {
      expect(Object.keys(theme.colors).sort()).toEqual([...SLOTS].sort());
      for (const value of Object.values(theme.colors)) {
        expect(value).toMatch(/^#[0-9a-f]{6}$/);
      }
      expect(theme.fonts.major.length).toBeGreaterThan(0);
      expect(theme.fonts.minor.length).toBeGreaterThan(0);
    }
  });

  it("pins each built-in's accent row (the gallery card swatches)", () => {
    const accents = (theme: DocumentTheme): string[] =>
      ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map(
        (slot) => theme.colors[slot as "accent1"],
      );
    expect(accents(BUILT_IN_THEMES[1])).toEqual(["#90c226", "#54a021", "#e6b91e", "#e76618", "#c42f1a", "#918655"]);
    expect(accents(BUILT_IN_THEMES[2])).toEqual(["#b01513", "#ea6312", "#e6b729", "#6aac90", "#5f9c9d", "#9b75b2"]);
    expect(accents(BUILT_IN_THEMES[3])).toEqual(["#e48312", "#bd582c", "#865640", "#9b8357", "#c2bc80", "#94a088"]);
    expect(accents(BUILT_IN_THEMES[4])).toEqual(["#1cade4", "#2683c6", "#27ced7", "#42ba97", "#3e8853", "#62a39f"]);
    expect(accents(BUILT_IN_THEMES[5])).toEqual(["#a53010", "#de7e18", "#9f8351", "#728653", "#92aa4c", "#6aac91"]);
  });
});

describe("resolveColor", () => {
  it("passes a literal through untouched", () => {
    expect(resolveColor("#123abc", DEFAULT_THEME)).toBe("#123abc");
  });

  it("resolves a plain reference to its slot", () => {
    expect(resolveColor({ theme: "accent1" }, DEFAULT_THEME)).toBe("#4472c4");
    expect(resolveColor({ theme: "dk2" }, DEFAULT_THEME)).toBe("#44546a");
  });

  it("applies lumMod + lumOff (PowerPoint Accent 1, Lighter 40% = #8faadc)", () => {
    expect(resolveColor({ theme: "accent1", lumMod: 0.6, lumOff: 0.4 }, DEFAULT_THEME)).toBe("#8faadc");
  });

  it("applies shade (c·val)", () => {
    expect(resolveColor({ theme: "lt1", shade: 0.5 }, DEFAULT_THEME)).toBe("#808080");
  });

  it("applies tint (c·val + (1−val))", () => {
    expect(resolveColor({ theme: "dk1", tint: 0.5 }, DEFAULT_THEME)).toBe("#808080");
  });

  it("applies lumMod alone (PowerPoint Darker 25%)", () => {
    // accent1 #4472c4 at lumMod 75% is PowerPoint's "Blue, Accent 1, Darker 25%".
    expect(resolveColor({ theme: "accent1", lumMod: 0.75 }, DEFAULT_THEME)).toBe("#2f5597");
  });
});

describe("resolveFont / guards", () => {
  it("resolves major/minor and passes a literal through", () => {
    expect(resolveFont({ theme: "major" }, DEFAULT_THEME)).toBe("Calibri Light");
    expect(resolveFont({ theme: "minor" }, DEFAULT_THEME)).toBe("Calibri");
    expect(resolveFont("Georgia", DEFAULT_THEME)).toBe("Georgia");
  });

  it("classifies references vs literals", () => {
    expect(isThemeColor({ theme: "accent1" })).toBe(true);
    expect(isThemeColor("#4472c4")).toBe(false);
    expect(isThemeColor(undefined)).toBe(false);
    expect(isThemeFont({ theme: "minor" })).toBe(true);
    expect(isThemeFont("Calibri")).toBe(false);
  });
});

describe("built-in contrast", () => {
  // WCAG relative luminance over the sRGB channels.
  const luminance = (hex: string): number => {
    const [r, g, b] = hexToRgb(hex).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  };
  const contrast = (a: string, b: string): number => {
    const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
    return (hi + 0.05) / (lo + 0.05);
  };

  it.each(BUILT_IN_THEMES.map((theme) => [theme.name, theme] as const))(
    "%s keeps its dark and light slots legible against each other",
    (_name, theme) => {
      expect(contrast(theme.colors.dk1, theme.colors.lt1)).toBeGreaterThanOrEqual(7);
      expect(contrast(theme.colors.dk2, theme.colors.lt2)).toBeGreaterThanOrEqual(3);
      // The dark slots really are the dark ones.
      expect(luminance(theme.colors.dk1)).toBeLessThan(luminance(theme.colors.lt1));
      expect(luminance(theme.colors.dk2)).toBeLessThan(luminance(theme.colors.lt2));
    },
  );
});

describe("picker swatch catalogue", () => {
  it("offers ten slots with five variants each", () => {
    expect(THEME_SWATCH_SLOTS).toEqual(["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]);
    expect(THEME_SWATCH_VARIANTS).toHaveLength(5);
    expect(THEME_SWATCH_VARIANTS.map((variant) => variant.label)).toEqual([
      "lighter 80%",
      "lighter 60%",
      "lighter 40%",
      "darker 25%",
      "darker 50%",
    ]);
  });
});
