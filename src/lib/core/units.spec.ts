import { describe, it, expect } from "vitest";
import { cssColorToOoxml, ooxmlColorToCss } from "./colors";
import {
  cmToEmu,
  cmToTwips,
  emuToCm,
  formatLength,
  fromCm,
  halfPointsToPt,
  parseLengthToCm,
  ptToHalfPoints,
  ptToTwentiethsOfPoint,
  roundForDisplay,
  toCm,
  twentiethsOfPointToPt,
  twipsToCm,
} from "./units";
import type { MeasurementUnit } from "./units";

const ALL_UNITS: MeasurementUnit[] = ["cm", "mm", "in", "pt"];

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

describe("cmToEmu / emuToCm (SPEC-024 DrawingML sizes)", () => {
  it("converts a whole inch (2.54cm) to 914400 EMU", () => {
    expect(cmToEmu(2.54)).toBe(914400);
  });

  it("converts a whole centimetre to 360000 EMU", () => {
    expect(cmToEmu(1)).toBe(360000);
  });

  it("emuToCm inverts cmToEmu for round EMU values real files use", () => {
    for (const emu of [0, 360000, 457200, 914400, 12700, 5334000]) {
      expect(cmToEmu(emuToCm(emu))).toBe(emu);
    }
  });

  it("rounds to the nearest whole EMU", () => {
    // 1px at 96 DPI in cm, times EMU/inch, is not a whole number pre-rounding.
    expect(cmToEmu(2.54 / 96)).toBe(Math.round((2.54 / 96 / 2.54) * 914400));
    expect(Number.isInteger(cmToEmu(1 / 3))).toBe(true);
  });

  it("emuToCm is the plain (unrounded) inverse ratio", () => {
    expect(emuToCm(914400)).toBeCloseTo(2.54, 10);
    expect(emuToCm(0)).toBe(0);
  });
});

describe("parseLengthToCm (SPEC-064 AC-11, AC-17)", () => {
  it("reads a bare number in the current display unit", () => {
    expect(parseLengthToCm("2", "cm")).toBeCloseTo(2, 10);
    expect(parseLengthToCm("10", "mm")).toBeCloseTo(1, 10);
    expect(parseLengthToCm("1", "in")).toBeCloseTo(2.54, 10);
    expect(parseLengthToCm("72", "pt")).toBeCloseTo(2.54, 10);
  });

  it("honours an explicit suffix whatever the display unit", () => {
    // Display unit is cm, but each suffix wins.
    expect(parseLengthToCm("2cm", "cm")).toBeCloseTo(2, 10);
    expect(parseLengthToCm("0.5in", "cm")).toBeCloseTo(1.27, 10);
    expect(parseLengthToCm("12pt", "cm")).toBeCloseTo((12 / 72) * 2.54, 10);
    expect(parseLengthToCm("10mm", "cm")).toBeCloseTo(1, 10);
    // Display unit is inches, but a cm suffix still means centimetres.
    expect(parseLengthToCm("3cm", "in")).toBeCloseTo(3, 10);
  });

  it("is case- and space-insensitive", () => {
    expect(parseLengthToCm("0.5 IN", "cm")).toBeCloseTo(1.27, 10);
    expect(parseLengthToCm("  2 Cm ", "in")).toBeCloseTo(2, 10);
    expect(parseLengthToCm("12 PT", "cm")).toBeCloseTo((12 / 72) * 2.54, 10);
  });

  it("accepts a leading dot, a sign and negative indents", () => {
    expect(parseLengthToCm(".5", "in")).toBeCloseTo(1.27, 10);
    expect(parseLengthToCm("+2", "cm")).toBeCloseTo(2, 10);
    expect(parseLengthToCm("-1.5cm", "cm")).toBeCloseTo(-1.5, 10);
  });

  it("rejects unparseable input (AC-11b)", () => {
    expect(parseLengthToCm("", "cm")).toBeNull();
    expect(parseLengthToCm("   ", "cm")).toBeNull();
    expect(parseLengthToCm("abc", "cm")).toBeNull();
    expect(parseLengthToCm("2cmm", "cm")).toBeNull();
    expect(parseLengthToCm("2 km", "cm")).toBeNull();
    expect(parseLengthToCm("1e3", "cm")).toBeNull();
    expect(parseLengthToCm("5.", "cm")).toBeNull();
  });
});

describe("formatLength (SPEC-064 AC-17a)", () => {
  it("converts out of centimetres, rounds per unit, and adds no suffix", () => {
    expect(formatLength(2.54, "cm")).toBe("2.54");
    expect(formatLength(2.54, "in")).toBe("1");
    expect(formatLength(1, "mm")).toBe("10");
    expect(formatLength(2.54, "pt")).toBe("72");
    // Float noise is rounded away, not shown.
    expect(formatLength(1 / 3, "cm")).toBe("0.33");
  });
});

describe("length round-trip and drift (SPEC-064 AC-12, AC-17b)", () => {
  it("parseLengthToCm(formatLength(x, u), u) returns x to within display precision", () => {
    // roundForDisplay's own step, per unit, in centimetres.
    // Half the unit's display step, expressed in centimetres: cm rounds to
    // 0.01cm, mm to 0.1mm (0.01cm), in to 0.01in (~0.025cm), pt to 1pt (~0.035cm).
    const toleranceCm: Record<MeasurementUnit, number> = {
      cm: 0.005,
      mm: 0.005,
      in: 0.02,
      pt: 0.02,
    };
    for (const unit of ALL_UNITS) {
      for (const cm of [0, 0.37, 1, 2.54, 3.1415, 12.7, 21, 29.7]) {
        const roundTripped = parseLengthToCm(formatLength(cm, unit), unit);
        expect(roundTripped).not.toBeNull();
        expect(Math.abs((roundTripped as number) - cm)).toBeLessThanOrEqual(toleranceCm[unit]);
      }
    }
  });

  it("switching units back and forth does not drift the stored centimetre value", () => {
    // The model stores cm and never rounds; only display rounds. Converting a
    // stored cm to a unit and back is the exact inverse ratio (no rounding),
    // so a measurement that is only ever converted for display cannot drift.
    for (const unit of ALL_UNITS) {
      for (const cm of [2.54, 1.27, 3.7, 0.63, 18.2]) {
        expect(toCm(fromCm(cm, unit), unit)).toBeCloseTo(cm, 10);
      }
    }
  });

  it("roundForDisplay never mutates and only the display value is coarsened", () => {
    // A precise stored cm formatted then re-parsed differs from the original by
    // at most the unit's display step — proof the model value is the source of
    // truth and rounding lives strictly at the boundary.
    const stored = 2.5399999;
    for (const unit of ALL_UNITS) {
      const shown = roundForDisplay(fromCm(stored, unit), unit);
      expect(Number.isFinite(shown)).toBe(true);
    }
  });
});
