import { describe, it, expect } from "vitest";
import {
  matchPaperSize,
  MARGINS_TOO_LARGE_MESSAGE,
  orientationOf,
  pageSizeFor,
  swapOrientation,
  textAreaTooNarrow,
} from "./pageSetup";
import { fromCm, roundForDisplay, toCm } from "./units";

describe("matchPaperSize (SPEC-027 AC-1, AC-8)", () => {
  it("matches A4 portrait", () => {
    expect(matchPaperSize({ widthCm: 21, heightCm: 29.7 })).toBe("A4");
  });

  it("matches A4 landscape (swapped dimensions)", () => {
    expect(matchPaperSize({ widthCm: 29.7, heightCm: 21 })).toBe("A4");
  });

  it("matches Letter within twip-rounding tolerance", () => {
    expect(matchPaperSize({ widthCm: 21.59, heightCm: 27.939999 })).toBe("Letter");
  });

  it("matches every standard size", () => {
    expect(matchPaperSize({ widthCm: 21.59, heightCm: 35.56 })).toBe("Legal");
    expect(matchPaperSize({ widthCm: 29.7, heightCm: 42 })).toBe("A3");
    expect(matchPaperSize({ widthCm: 14.8, heightCm: 21 })).toBe("A5");
    expect(matchPaperSize({ widthCm: 27.94, heightCm: 43.18 })).toBe("Tabloid");
  });

  it("falls back to Custom for a non-standard size (AC-8)", () => {
    expect(matchPaperSize({ widthCm: 12, heightCm: 18 })).toBe("Custom");
  });
});

describe("orientationOf / swapOrientation (SPEC-027 AC-2)", () => {
  it("reads portrait when width <= height", () => {
    expect(orientationOf({ widthCm: 21, heightCm: 29.7 })).toBe("portrait");
  });

  it("reads landscape when width > height", () => {
    expect(orientationOf({ widthCm: 29.7, heightCm: 21 })).toBe("landscape");
  });

  it("swaps width and height", () => {
    expect(swapOrientation({ widthCm: 21, heightCm: 29.7 })).toEqual({ widthCm: 29.7, heightCm: 21 });
  });
});

describe("pageSizeFor", () => {
  it("returns the portrait catalogue entry unchanged", () => {
    expect(pageSizeFor("A4", "portrait")).toEqual({ widthCm: 21, heightCm: 29.7 });
  });

  it("swaps for landscape", () => {
    expect(pageSizeFor("A4", "landscape")).toEqual({ widthCm: 29.7, heightCm: 21 });
  });
});

describe("unit conversion (SPEC-027 AC-4)", () => {
  it("round-trips cm exactly", () => {
    expect(toCm(2.5, "cm")).toBe(2.5);
    expect(fromCm(2.5, "cm")).toBe(2.5);
  });

  it("converts mm", () => {
    expect(toCm(25, "mm")).toBeCloseTo(2.5, 10);
    expect(fromCm(2.5, "mm")).toBeCloseTo(25, 10);
  });

  it("converts inches", () => {
    expect(toCm(1, "in")).toBeCloseTo(2.54, 10);
    expect(fromCm(2.54, "in")).toBeCloseTo(1, 10);
  });

  it("converts points (72pt = 1in)", () => {
    expect(toCm(72, "pt")).toBeCloseTo(2.54, 10);
    expect(fromCm(2.54, "pt")).toBeCloseTo(72, 10);
  });

  it("rounds display values to a sensible precision per unit", () => {
    expect(roundForDisplay(2.545, "cm")).toBe(2.55);
    expect(roundForDisplay(25.449, "mm")).toBe(25.4);
    expect(roundForDisplay(28.3, "pt")).toBe(28);
  });
});

describe("textAreaTooNarrow (SPEC-027 AC-9)", () => {
  const a4 = { widthCm: 21, heightCm: 29.7 };

  it("accepts standard margins", () => {
    expect(textAreaTooNarrow(a4, { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 })).toBe(false);
  });

  it("rejects margins that leave under 1cm of horizontal text area", () => {
    expect(textAreaTooNarrow(a4, { topCm: 2, rightCm: 10, bottomCm: 2, leftCm: 10.5 })).toBe(true);
  });

  it("rejects margins that leave under 1cm of vertical text area", () => {
    expect(textAreaTooNarrow(a4, { topCm: 14.5, rightCm: 2, bottomCm: 14.5, leftCm: 2 })).toBe(true);
  });

  it("accepts a text area of exactly 1cm", () => {
    expect(textAreaTooNarrow({ widthCm: 3, heightCm: 29.7 }, { topCm: 0, rightCm: 1, bottomCm: 0, leftCm: 1 })).toBe(
      false,
    );
  });

  it("exports the exact rejection message text (AC-9)", () => {
    expect(MARGINS_TOO_LARGE_MESSAGE).toBe("Margins are too large for this page size");
  });
});
