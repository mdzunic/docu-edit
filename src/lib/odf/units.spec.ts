import { describe, it, expect } from "vitest";
import { odfColorToCss, odfLengthToCm, odfLengthToPt, odfPercentage } from "./units";

describe("odfLengthToCm", () => {
  it("parses centimetre lengths", () => {
    expect(odfLengthToCm("2.54cm")).toBeCloseTo(2.54, 6);
  });

  it("parses millimetre lengths", () => {
    expect(odfLengthToCm("10mm")).toBeCloseTo(1, 6);
  });

  it("parses inch lengths", () => {
    expect(odfLengthToCm("1in")).toBeCloseTo(2.54, 6);
  });

  it("parses point and pica lengths", () => {
    expect(odfLengthToCm("72pt")).toBeCloseTo(2.54, 6);
    expect(odfLengthToCm("6pc")).toBeCloseTo(2.54, 6);
  });

  it("parses negative lengths (hanging indents)", () => {
    expect(odfLengthToCm("-0.25in")).toBeCloseTo(-0.635, 6);
  });

  it("returns undefined for percentages, unitless and missing values", () => {
    expect(odfLengthToCm("115%")).toBeUndefined();
    expect(odfLengthToCm("12")).toBeUndefined();
    expect(odfLengthToCm("")).toBeUndefined();
    expect(odfLengthToCm(null)).toBeUndefined();
    expect(odfLengthToCm(undefined)).toBeUndefined();
  });
});

describe("odfLengthToPt", () => {
  it("converts absolute lengths to points", () => {
    expect(odfLengthToPt("12pt")).toBeCloseTo(12, 6);
    expect(odfLengthToPt("1in")).toBeCloseTo(72, 6);
    expect(odfLengthToPt("0.423cm")).toBeCloseTo(11.99, 2);
  });

  it("returns undefined for non-lengths", () => {
    expect(odfLengthToPt("130%")).toBeUndefined();
    expect(odfLengthToPt("normal")).toBeUndefined();
  });
});

describe("odfPercentage", () => {
  it("parses percentages", () => {
    expect(odfPercentage("130%")).toBe(130);
    expect(odfPercentage("115.5%")).toBe(115.5);
  });

  it("returns undefined for lengths and missing values", () => {
    expect(odfPercentage("12pt")).toBeUndefined();
    expect(odfPercentage(null)).toBeUndefined();
  });
});

describe("odfColorToCss", () => {
  it("lowercases hex colours", () => {
    expect(odfColorToCss("#C9211E")).toBe("#c9211e");
  });

  it("returns undefined for transparent/keyword/missing values", () => {
    expect(odfColorToCss("transparent")).toBeUndefined();
    expect(odfColorToCss("")).toBeUndefined();
    expect(odfColorToCss(null)).toBeUndefined();
  });
});
