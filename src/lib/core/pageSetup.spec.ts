import { describe, it, expect } from "vitest";
import {
  matchPaperSize,
  matchSlideSize,
  MARGINS_TOO_LARGE_MESSAGE,
  MAX_SLIDE_DIMENSION_CM,
  MIN_SLIDE_DIMENSION_CM,
  orientationOf,
  pageSizeFor,
  slideSizeFor,
  SLIDE_SIZES,
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

describe("matchSlideSize (SPEC-090)", () => {
  it("matches each catalogue size in landscape", () => {
    expect(matchSlideSize({ widthCm: 33.867, heightCm: 19.05 })).toBe("16:9");
    expect(matchSlideSize({ widthCm: 25.4, heightCm: 19.05 })).toBe("4:3");
    expect(matchSlideSize({ widthCm: 25.4, heightCm: 15.875 })).toBe("16:10");
    expect(matchSlideSize(SLIDE_SIZES.A4)).toBe("A4");
    expect(matchSlideSize(SLIDE_SIZES.Letter)).toBe("Letter");
  });

  it("matches in the swapped (portrait) orientation too", () => {
    expect(matchSlideSize({ widthCm: 19.05, heightCm: 33.867 })).toBe("16:9");
    expect(matchSlideSize(swapOrientation(SLIDE_SIZES.A4))).toBe("A4");
  });

  it("tolerates up to ±0.05 cm and rejects beyond it", () => {
    expect(matchSlideSize({ widthCm: 33.9, heightCm: 19.02 })).toBe("16:9");
    expect(matchSlideSize({ widthCm: 34.0, heightCm: 19.05 })).toBe("Custom");
  });

  it("returns Custom for a non-standard size", () => {
    expect(matchSlideSize({ widthCm: 20, heightCm: 20 })).toBe("Custom");
  });

  it("prefers 4:3 over 16:10 when both could tie only by one axis (distinct heights)", () => {
    // 4:3 and 16:10 share width 25.4 but differ in height; the exact size wins.
    expect(matchSlideSize({ widthCm: 25.4, heightCm: 19.05 })).toBe("4:3");
    expect(matchSlideSize({ widthCm: 25.4, heightCm: 15.875 })).toBe("16:10");
  });
});

describe("slideSizeFor (SPEC-090)", () => {
  it("returns the landscape catalogue entry by default", () => {
    expect(slideSizeFor("16:9", "landscape")).toEqual({ widthCm: 33.867, heightCm: 19.05 });
    expect(slideSizeFor("4:3", "landscape")).toEqual({ widthCm: 25.4, heightCm: 19.05 });
  });

  it("swaps to portrait when requested", () => {
    expect(slideSizeFor("16:9", "portrait")).toEqual({ widthCm: 19.05, heightCm: 33.867 });
  });

  it("A4/Letter reuse PAPER_SIZES swapped to landscape", () => {
    expect(slideSizeFor("A4", "landscape")).toEqual({ widthCm: 29.7, heightCm: 21 });
    expect(slideSizeFor("Letter", "landscape")).toEqual({ widthCm: 27.94, heightCm: 21.59 });
  });
});

describe("slide dimension clamps (SPEC-090)", () => {
  it("are 1in and 56in in centimetres", () => {
    expect(MIN_SLIDE_DIMENSION_CM).toBe(2.54);
    expect(MAX_SLIDE_DIMENSION_CM).toBe(142.24);
  });
});
