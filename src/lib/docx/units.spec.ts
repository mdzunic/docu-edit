import { describe, it, expect } from "vitest";
import {
  cmToTwips,
  cssColorToOoxml,
  halfPointsToPt,
  ooxmlColorToCss,
  ptToHalfPoints,
  ptToTwentiethsOfPoint,
  twentiethsOfPointToPt,
  twipsToCm,
} from "./units";

describe("inverse unit conversions (SPEC-006 writer)", () => {
  it("cmToTwips inverts twipsToCm exactly for the integers real files use", () => {
    // Page sizes, margins and indents from the SPEC-004 fixtures.
    for (const twips of [0, 360, 708, 720, 1134, 1440, 11906, 12240, 15840, 16838]) {
      expect(cmToTwips(twipsToCm(twips))).toBe(twips);
    }
  });

  it("cmToTwips converts whole centimetres", () => {
    expect(cmToTwips(2.54)).toBe(1440);
    expect(cmToTwips(1.27)).toBe(720);
  });

  it("ptToHalfPoints inverts halfPointsToPt", () => {
    for (const halfPoints of [22, 26, 28, 32, 40, 21]) {
      expect(ptToHalfPoints(halfPointsToPt(halfPoints))).toBe(halfPoints);
    }
    expect(ptToHalfPoints(11)).toBe(22);
    expect(ptToHalfPoints(10.5)).toBe(21);
  });

  it("ptToTwentiethsOfPoint inverts twentiethsOfPointToPt", () => {
    for (const value of [0, 160, 240, 360]) {
      expect(ptToTwentiethsOfPoint(twentiethsOfPointToPt(value))).toBe(value);
    }
    expect(ptToTwentiethsOfPoint(12)).toBe(240);
  });

  it("cssColorToOoxml inverts ooxmlColorToCss", () => {
    expect(cssColorToOoxml(ooxmlColorToCss("1155CC"))).toBe("1155CC");
    expect(cssColorToOoxml("#ff0000")).toBe("FF0000");
  });

  it("cssColorToOoxml rejects anything but six-digit hex", () => {
    expect(cssColorToOoxml(undefined)).toBeUndefined();
    expect(cssColorToOoxml("")).toBeUndefined();
    expect(cssColorToOoxml("red")).toBeUndefined();
    expect(cssColorToOoxml("#fff")).toBeUndefined();
    expect(cssColorToOoxml("ff0000")).toBeUndefined();
  });
});
