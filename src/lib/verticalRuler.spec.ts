import { describe, it, expect } from "vitest";
import {
  valueFromVerticalPosition,
  verticalMarkerPositionCm,
  rulerTicks,
  rulerTickScheme,
  type VerticalRulerContext,
} from "./verticalRuler";
import { toCm } from "./core/units";

// A4 portrait with Word's default 2.54cm margins and 1.27cm bands.
function ctx(overrides: Partial<VerticalRulerContext> = {}): VerticalRulerContext {
  return {
    pageHeightCm: 29.7,
    marginTopCm: 2.54,
    marginBottomCm: 2.54,
    headerDistanceCm: 1.27,
    footerDistanceCm: 1.27,
    ...overrides,
  };
}

describe("rulerTickScheme (SPEC-064 Tick scheme)", () => {
  it("labels 1cm majors and draws 0.5cm minors for cm at 100%", () => {
    expect(rulerTickScheme("cm", 1)).toEqual({ majorCm: 1, minorCm: 0.5 });
  });

  it("draws only major ticks below 50% zoom", () => {
    expect(rulerTickScheme("cm", 0.4)).toEqual({ majorCm: 1, minorCm: 1 });
    expect(rulerTickScheme("in", 0.4).minorCm).toBeCloseTo(toCm(1, "in"), 10);
  });

  it("adds the next subdivision above 150% zoom", () => {
    expect(rulerTickScheme("cm", 2).minorCm).toBeCloseTo(0.25, 10);
    expect(rulerTickScheme("in", 2).minorCm).toBeCloseTo(toCm(1 / 16, "in"), 10);
  });

  it("coarsens inches to quarter-inch minors below 75% zoom", () => {
    expect(rulerTickScheme("in", 0.6).minorCm).toBeCloseTo(toCm(1 / 4, "in"), 10);
    expect(rulerTickScheme("in", 1).minorCm).toBeCloseTo(toCm(1 / 8, "in"), 10);
  });

  it("labels 72pt (one inch) majors and 36pt minors for pt", () => {
    const scheme = rulerTickScheme("pt", 1);
    expect(scheme.majorCm).toBeCloseTo(toCm(72, "pt"), 10);
    expect(scheme.minorCm).toBeCloseTo(toCm(36, "pt"), 10);
  });
});

describe("rulerTicks (SPEC-064 AC-4, AC-4a)", () => {
  it("numbers outward from the top of the text area, labelling the origin 0", () => {
    const ticks = rulerTicks(29.7, 2.54, "cm", 1);
    const origin = ticks.find((t) => Math.abs(t.cm - 2.54) < 1e-6);
    expect(origin?.major).toBe(true);
    expect(origin?.label).toBe("0");
    // A major tick 3cm below the origin reads "3".
    const three = ticks.find((t) => Math.abs(t.cm - (2.54 + 3)) < 1e-6);
    expect(three?.label).toBe("3");
    // A major tick 2cm ABOVE the origin (inside the top margin) also reads "2".
    const twoAbove = ticks.find((t) => Math.abs(t.cm - (2.54 - 2)) < 1e-6);
    expect(twoAbove?.label).toBe("2");
  });

  it("marks half-centimetre minors unlabelled at 100%", () => {
    const ticks = rulerTicks(29.7, 2.54, "cm", 1);
    const half = ticks.find((t) => Math.abs(t.cm - (2.54 + 0.5)) < 1e-6);
    expect(half?.major).toBe(false);
    expect(half?.label).toBe("");
  });

  it("keeps every tick inside the page", () => {
    const ticks = rulerTicks(29.7, 2.54, "cm", 1);
    for (const tick of ticks) {
      expect(tick.cm).toBeGreaterThanOrEqual(-1e-6);
      expect(tick.cm).toBeLessThanOrEqual(29.7 + 1e-6);
    }
  });

  it("draws fewer ticks below 50% zoom (majors only)", () => {
    const majorsOnly = rulerTicks(29.7, 2.54, "cm", 0.4);
    expect(majorsOnly.every((t) => t.major)).toBe(true);
  });

  it("returns nothing for a non-positive page height", () => {
    expect(rulerTicks(0, 0, "cm", 1)).toEqual([]);
  });
});

describe("verticalMarkerPositionCm (SPEC-064 AC-2, AC-3)", () => {
  it("places margins and bands relative to the correct edge", () => {
    expect(verticalMarkerPositionCm("margin-top", ctx())).toBeCloseTo(2.54, 10);
    expect(verticalMarkerPositionCm("margin-bottom", ctx())).toBeCloseTo(29.7 - 2.54, 10);
    expect(verticalMarkerPositionCm("header", ctx())).toBeCloseTo(1.27, 10);
    expect(verticalMarkerPositionCm("footer", ctx())).toBeCloseTo(29.7 - 1.27, 10);
  });
});

describe("valueFromVerticalPosition (SPEC-064 AC-2a, AC-3a/3b)", () => {
  it("maps a top-margin drag straight to marginTopCm", () => {
    expect(valueFromVerticalPosition("margin-top", 3, ctx())).toEqual({ kind: "marginTopCm", valueCm: 3 });
  });

  it("maps a bottom-margin drag to distance from the bottom edge", () => {
    const update = valueFromVerticalPosition("margin-bottom", 29.7 - 3, ctx());
    expect(update).toEqual({ kind: "marginBottomCm", valueCm: 3 });
  });

  it("maps header/footer band drags to their distances", () => {
    expect(valueFromVerticalPosition("header", 1.5, ctx())).toEqual({ kind: "headerDistanceCm", valueCm: 1.5 });
    const footer = valueFromVerticalPosition("footer", 29.7 - 0.8, ctx());
    expect(footer.kind).toBe("footerDistanceCm");
    expect(footer.valueCm).toBeCloseTo(0.8, 10);
  });

  it("clamps the top margin so it cannot cross the bottom margin", () => {
    const update = valueFromVerticalPosition("margin-top", 40, ctx());
    expect(update.valueCm).toBeCloseTo(29.7 - 2.54, 10);
  });

  it("snaps a drag to the 0.05cm grid", () => {
    const update = valueFromVerticalPosition("margin-top", 3.013, ctx());
    expect(update.valueCm).toBeCloseTo(3.0, 10);
  });

  it("clamps a negative drag to zero", () => {
    expect(valueFromVerticalPosition("header", -5, ctx()).valueCm).toBe(0);
  });
});
