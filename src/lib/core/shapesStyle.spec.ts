// SPEC-073: the widened styling model's pure core — lighterTint (D2),
// normaliseShapeFill (Migration/AC-30), the five style presets (D3/AC-19),
// the shadow presets (D4/AC-21) and the seeded gradient (AC-22).
import { describe, it, expect } from "vitest";
import {
  DEFAULT_NEW_SHAPE_OUTLINE,
  DEFAULT_SHAPE_SHADOW,
  SHAPE_SHADOW_PRESETS,
  SHAPE_STYLE_PRESETS,
  lighterTint,
  normaliseShapeFill,
  seededGradient,
  shapeFillBaseColor,
  shapeFillIsVisible,
  type ShapeFill,
} from "./shapes";
import { DEFAULT_THEME, resolveColor } from "./theme";

describe("lighterTint (D2, AC-23)", () => {
  it("tints the accent #4472c4 to exactly #8faadc", () => {
    expect(lighterTint("#4472c4")).toBe("#8faadc");
  });
  it("tints black to #666666", () => {
    expect(lighterTint("#000000")).toBe("#666666");
  });
  it("leaves white white", () => {
    expect(lighterTint("#ffffff")).toBe("#ffffff");
  });
  it("lightens a mid colour (luminance strictly increases)", () => {
    const tinted = lighterTint("#70ad47");
    expect(tinted).not.toBe("#70ad47");
    const lum = (hex: string) =>
      parseInt(hex.slice(1, 3), 16) + parseInt(hex.slice(3, 5), 16) + parseInt(hex.slice(5, 7), 16);
    expect(lum(tinted)).toBeGreaterThan(lum("#70ad47"));
  });
});

describe("normaliseShapeFill (AC-30) — all four cases", () => {
  it("migrates the legacy #rrggbb string to a solid fill", () => {
    expect(normaliseShapeFill("#4472c4")).toEqual({ kind: "solid", color: "#4472c4" });
  });
  it("returns a union value as-is (same reference)", () => {
    const fill: ShapeFill = { kind: "gradient", gradient: { type: "linear", angleDeg: 90, stops: [{ position: 0, color: "#000000" }, { position: 1, color: "#ffffff" }] } };
    expect(normaliseShapeFill(fill)).toBe(fill);
  });
  it("maps undefined to undefined", () => {
    expect(normaliseShapeFill(undefined)).toBeUndefined();
  });
  it("maps anything else to undefined", () => {
    expect(normaliseShapeFill("red")).toBeUndefined();
    expect(normaliseShapeFill(42)).toBeUndefined();
    expect(normaliseShapeFill({ some: "junk" })).toBeUndefined();
  });
});

describe("the five style presets (D3, AC-18/AC-19)", () => {
  it("is exactly the five ids in gallery order", () => {
    expect(SHAPE_STYLE_PRESETS.map((p) => p.id)).toEqual([
      "filled",
      "subtleFilled",
      "outlined",
      "outlinedShadow",
      "noFill",
    ]);
  });
  // SPEC-077 AC-43: the colours SPEC-073 drew from the accent — the base
  // #4472c4 and its lighter variants — are theme references to accent1, so a
  // theme change restyles shapes the presets coloured. AC-44 pins each
  // reference's Office-theme resolution to the exact values SPEC-073 shipped.
  // The filled preset's border stays the shipped literal #2e5395: it was
  // never the accent or a lighter variant of it, and no canonical
  // lumMod/lumOff/shade/tint combination reproduces it from accent1
  // (verified by exhaustive search over the transform grid; the nearest,
  // Darker 25%, resolves to #2f5597), so only a literal keeps AC-44 exact.
  it("filled is an accent1 fill with the shipped default border (AC-43/AC-44)", () => {
    const filled = SHAPE_STYLE_PRESETS[0];
    expect(filled.fill).toEqual({ kind: "solid", color: { theme: "accent1" } });
    expect(filled.outline).toEqual(DEFAULT_NEW_SHAPE_OUTLINE);
    expect(filled.shadow).toBeUndefined();
  });
  it("subtleFilled is accent1 at Lighter 80% with an accent1 1pt border (refs, AC-43)", () => {
    const subtle = SHAPE_STYLE_PRESETS[1];
    expect(subtle.fill).toEqual({ kind: "solid", color: { theme: "accent1", lumMod: 0.2, lumOff: 0.8 } });
    expect(subtle.outline).toEqual({ color: { theme: "accent1" }, widthPt: 1, dash: "solid" });
  });
  it("outlined is no fill + accent1/1.5pt/solid (refs, AC-43)", () => {
    const outlined = SHAPE_STYLE_PRESETS[2];
    expect(outlined.fill).toBeUndefined();
    expect(outlined.outline).toEqual({ color: { theme: "accent1" }, widthPt: 1.5, dash: "solid" });
    expect(outlined.shadow).toBeUndefined();
  });
  it("outlinedShadow is outlined plus the D4 standard shadow", () => {
    const preset = SHAPE_STYLE_PRESETS[3];
    expect(preset.outline).toEqual({ color: { theme: "accent1" }, widthPt: 1.5, dash: "solid" });
    expect(preset.shadow).toEqual({ color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 });
  });
  it("resolves to SPEC-073's shipped values under the Office theme (AC-44)", () => {
    const [filled, subtle, outlined, outlinedShadow] = SHAPE_STYLE_PRESETS;
    const resolveSolid = (fill: ShapeFill | undefined) =>
      fill?.kind === "solid" ? resolveColor(fill.color, DEFAULT_THEME) : undefined;
    // DEFAULT_NEW_SHAPE_FILL's shipped solid colour.
    expect(resolveSolid(filled.fill)).toBe("#4472c4");
    expect(resolveColor(filled.outline!.color, DEFAULT_THEME)).toBe("#2e5395");
    expect(resolveSolid(subtle.fill)).toBe("#dae3f3");
    expect(resolveColor(subtle.outline!.color, DEFAULT_THEME)).toBe("#4472c4");
    expect(resolveColor(outlined.outline!.color, DEFAULT_THEME)).toBe("#4472c4");
    expect(resolveColor(outlinedShadow.outline!.color, DEFAULT_THEME)).toBe("#4472c4");
  });
  it("noFill clears everything", () => {
    const noFill = SHAPE_STYLE_PRESETS[4];
    expect(noFill.fill).toBeUndefined();
    expect(noFill.outline).toBeUndefined();
    expect(noFill.shadow).toBeUndefined();
  });
});

describe("shadow presets (D4, AC-20/AC-21)", () => {
  it("the standard shadow is D4's exact values", () => {
    expect(DEFAULT_SHAPE_SHADOW).toEqual({ color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 });
  });
  it("offers nine entries: No shadow plus the eight directions, identical except angleDeg", () => {
    expect(SHAPE_SHADOW_PRESETS).toHaveLength(9);
    expect(SHAPE_SHADOW_PRESETS[0].shadow).toBeUndefined();
    const angles = SHAPE_SHADOW_PRESETS.slice(1).map((p) => p.shadow!.angleDeg);
    expect([...angles].sort((a, b) => a - b)).toEqual([0, 45, 90, 135, 180, 225, 270, 315]);
    for (const preset of SHAPE_SHADOW_PRESETS.slice(1)) {
      expect({ ...preset.shadow!, angleDeg: 90 }).toEqual(DEFAULT_SHAPE_SHADOW);
    }
  });
});

describe("seededGradient (AC-22)", () => {
  it("seeds two linear stops at 90 degrees, base to lighterTint(base)", () => {
    expect(seededGradient("#4472c4")).toEqual({
      type: "linear",
      angleDeg: 90,
      stops: [
        { position: 0, color: "#4472c4" },
        { position: 1, color: "#8faadc" },
      ],
    });
  });
});

describe("fill inspection helpers", () => {
  it("shapeFillIsVisible: none and absent are invisible, the rest visible", () => {
    expect(shapeFillIsVisible(undefined)).toBe(false);
    expect(shapeFillIsVisible({ kind: "none" })).toBe(false);
    expect(shapeFillIsVisible({ kind: "solid", color: "#000000" })).toBe(true);
  });
  it("shapeFillBaseColor: solid colour, first gradient stop, pattern foreground", () => {
    expect(shapeFillBaseColor({ kind: "solid", color: "#123456" })).toBe("#123456");
    expect(
      shapeFillBaseColor({
        kind: "gradient",
        gradient: { type: "linear", stops: [{ position: 0, color: "#111111" }, { position: 1, color: "#222222" }] },
      }),
    ).toBe("#111111");
    expect(shapeFillBaseColor({ kind: "pattern", preset: "ltHorz", foreground: "#333333", background: "#ffffff" })).toBe("#333333");
    expect(shapeFillBaseColor({ kind: "none" })).toBeUndefined();
  });
});
