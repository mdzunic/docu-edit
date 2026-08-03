import { describe, it, expect } from "vitest";
import {
  cmToOdfLength,
  multiplierToOdfPercentage,
  odfColorToCss,
  odfLengthToCm,
  odfLengthToPt,
  odfPercentage,
  ptToOdfLength,
} from "./units";

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

describe("cmToOdfLength", () => {
  it("prefers the shortest decimal", () => {
    expect(cmToOdfLength(2)).toBe("2cm");
    expect(cmToOdfLength(1.27)).toBe("1.27cm");
    expect(cmToOdfLength(-0.635)).toBe("-0.635cm");
  });

  it("parses back to exactly the input for every source unit", () => {
    for (const source of ["1.27cm", "0.5in", "-0.25in", "21.001cm", "11in", "36pt", "10mm"]) {
      const value = odfLengthToCm(source)!;
      expect(odfLengthToCm(cmToOdfLength(value))).toBe(value);
    }
  });
});

describe("ptToOdfLength", () => {
  it("prefers the shortest decimal", () => {
    expect(ptToOdfLength(12)).toBe("12pt");
    expect(ptToOdfLength(0)).toBe("0pt");
    expect(ptToOdfLength(10.5)).toBe("10.5pt");
  });

  it("parses back to exactly the input for cm-denominated sources (LibreOffice spacing)", () => {
    for (const source of ["0.423cm", "0.212cm", "0.247cm", "0.199cm", "1in"]) {
      const value = odfLengthToPt(source)!;
      expect(odfLengthToPt(ptToOdfLength(value))).toBe(value);
    }
  });

  it("parses back to exactly the input for percentage-derived font sizes", () => {
    // 130% and 115% of a 14pt parent — the fixtures' heading sizes; 16.1
    // needs a decimal one ulp off the value itself.
    for (const value of [(14 * 130) / 100, (14 * 115) / 100]) {
      expect(odfLengthToPt(ptToOdfLength(value))).toBe(value);
    }
  });
});

describe("multiplierToOdfPercentage", () => {
  it("emits the shortest percentage that parses back exactly", () => {
    expect(multiplierToOdfPercentage(1.15)).toBe("115%");
    expect(multiplierToOdfPercentage(1.5)).toBe("150%");
    const value = odfPercentage(multiplierToOdfPercentage(1.15))! / 100;
    expect(value).toBe(1.15);
  });
});
