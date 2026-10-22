import { describe, expect, it } from "vitest";
import { emuToCm, emuToPt, lengthToEmu } from "./length";
import { DocumentGenerationError } from "./errors";

describe("lengthToEmu (AC-25)", () => {
  it("converts cm to the nearest whole Emu", () => {
    expect(lengthToEmu({ cm: 2.54 }, "width")).toBe(914400);
  });

  it("converts pt to the nearest whole Emu", () => {
    expect(lengthToEmu({ pt: 12 }, "width")).toBe(152400);
  });

  it("converts in to the nearest whole Emu", () => {
    expect(lengthToEmu({ in: 1 }, "width")).toBe(914400);
  });

  it("passes an integer emu value through unchanged", () => {
    expect(lengthToEmu({ emu: 12345 }, "width")).toBe(12345);
  });

  it("rounds a fractional cm/pt/in result to the nearest whole Emu", () => {
    expect(lengthToEmu({ cm: 1 }, "width")).toBe(360000);
    expect(Number.isInteger(lengthToEmu({ cm: 0.001 }, "width"))).toBe(true);
  });

  it("rejects a bare number", () => {
    expect(() => lengthToEmu(5, "width")).toThrow(DocumentGenerationError);
    expect(() => lengthToEmu(5, "width")).toThrow(/width/);
  });

  it("rejects an object with more than one unit key", () => {
    expect(() => lengthToEmu({ cm: 1, pt: 1 }, "width")).toThrow(DocumentGenerationError);
  });

  it("rejects an object with zero unit keys", () => {
    expect(() => lengthToEmu({}, "width")).toThrow(DocumentGenerationError);
  });

  it("rejects an unrecognised unit key", () => {
    expect(() => lengthToEmu({ mm: 5 }, "width")).toThrow(DocumentGenerationError);
  });

  it("rejects a non-integer emu value (AC-26)", () => {
    expect(() => lengthToEmu({ emu: 1.5 }, "width")).toThrow(/width\.emu/);
  });

  it("rejects a negative value for every unit (AC-26)", () => {
    expect(() => lengthToEmu({ cm: -1 }, "width")).toThrow(/non-negative/);
    expect(() => lengthToEmu({ pt: -1 }, "width")).toThrow(/non-negative/);
    expect(() => lengthToEmu({ in: -1 }, "width")).toThrow(/non-negative/);
    expect(() => lengthToEmu({ emu: -1 }, "width")).toThrow(/non-negative/);
  });

  it("rejects a non-finite value", () => {
    expect(() => lengthToEmu({ cm: NaN }, "width")).toThrow(DocumentGenerationError);
    expect(() => lengthToEmu({ cm: Infinity }, "width")).toThrow(DocumentGenerationError);
  });

  it("names the offending path in the error message (AC-7)", () => {
    expect(() => lengthToEmu({ cm: -1 }, "body[2].width")).toThrow(/body\[2\]\.width\.cm/);
  });
});

describe("emuToCm / emuToPt", () => {
  it("is the inverse of the cm/pt conversions", () => {
    expect(emuToCm(914400)).toBeCloseTo(2.54, 10);
    expect(emuToPt(152400)).toBeCloseTo(12, 10);
  });
});
