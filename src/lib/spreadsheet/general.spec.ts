// SPEC-116 "The General display": generalText shows a cell value as Excel's
// General format does at full column width — "" for no value, a string as
// itself, TRUE/FALSE, an error's code, and a number in at most 11
// characters: for 1e-9 ≤ |n| < 1e11 the fixed-notation text of n at up to 15
// significant digits, trailing zeros and point removed, decimals rounded off
// until it fits; otherwise scientific, the mantissa rounded so the whole
// text fits, an upper-case E and a signed exponent of at least two digits.
import { describe, expect, it } from "vitest";
import { generalText } from "./general";
import type { ErrorCode } from "./model";

/** The General format's width at a standard column. */
const WIDTH = 11;

/** Scientific General text: a normalised mantissa without trailing zeros,
 * upper-case E, a signed exponent of at least two digits. */
const SCIENTIFIC = /^-?[1-9](\.\d*[1-9])?E[+-]\d{2,}$/;

/** Fixed General text: no exponent, no leading zeros, no trailing zeros or
 * trailing point. */
const FIXED = /^-?(0|[1-9]\d*)(\.\d*[1-9])?$/;

/** The most a General text may differ from its number when it keeps as many
 * digits as 11 characters leave room for: half a unit in the last place
 * that fits (the spec's "round off decimals until it fits" / "mantissa
 * rounded so the whole text is at most 11 characters"). */
function maxRoundingError(text: string): number {
  const sign = text.startsWith("-") ? 1 : 0;
  const body = text.slice(sign);
  const e = body.indexOf("E");
  if (e >= 0) {
    // "d." + decimals + "E±XX"
    const decimals = WIDTH - sign - 2 - (body.length - e);
    return 0.5 * 10 ** (Number(body.slice(e + 1)) - decimals);
  }
  const integerDigits = body.split(".")[0]!.length;
  return 0.5 * 10 ** -(WIDTH - sign - integerDigits - 1);
}

/** 560 numbers from 1e-12 to about 1e16, both signs, with mantissas that
 * exercise rounding (thirds, sevenths, runs of 9s, 15 significant digits). */
const MANTISSAS = [1, 1.5, 2.5, Math.E, Math.PI, 10 / 7, 9.87654321, 9.999999999999, 1.23456789012345, 5.55555555555];
const SAMPLE = Array.from({ length: 28 }, (_, i) => i - 12).flatMap((exponent) =>
  MANTISSAS.flatMap((mantissa) => [mantissa * 10 ** exponent, -mantissa * 10 ** exponent]),
);

describe("generalText — values that are not numbers", () => {
  it("shows no value as empty text", () => {
    expect(generalText(null)).toBe("");
  });

  it("shows a string as itself, untrimmed and unparsed", () => {
    expect(generalText("text")).toBe("text");
    expect(generalText("")).toBe("");
    expect(generalText("007")).toBe("007");
    expect(generalText("TRUE")).toBe("TRUE");
    expect(generalText("  padded  ")).toBe("  padded  ");
    expect(generalText("two\nlines")).toBe("two\nlines");
  });

  it("shows booleans as TRUE and FALSE", () => {
    expect(generalText(true)).toBe("TRUE");
    expect(generalText(false)).toBe("FALSE");
  });

  it("shows an error as its code", () => {
    expect(generalText({ error: "#DIV/0!" })).toBe("#DIV/0!");
    const codes: ErrorCode[] = [
      "#NULL!",
      "#DIV/0!",
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#NUM!",
      "#N/A",
      "#GETTING_DATA",
      "#SPILL!",
      "#CALC!",
      "#FIELD!",
      "#BLOCKED!",
      "#CONNECT!",
      "#BUSY!",
      "#UNKNOWN!",
    ];
    for (const error of codes) {
      expect(generalText({ error })).toBe(error);
    }
  });
});

describe("generalText — numbers", () => {
  it.each<[label: string, value: number, text: string]>([
    ["0", 0, "0"],
    ["-0", -0, "0"],
    ["1", 1, "1"],
    ["-1", -1, "-1"],
    ["100", 100, "100"],
    ["0.5", 0.5, "0.5"],
    ["-0.5", -0.5, "-0.5"],
    ["1234.5678", 1234.5678, "1234.5678"],
    ["0.1 + 0.2", 0.1 + 0.2, "0.3"],
    ["1/3", 1 / 3, "0.333333333"],
    ["2/3", 2 / 3, "0.666666667"],
    ["-1/3", -1 / 3, "-0.33333333"],
    ["π", Math.PI, "3.141592654"],
    ["-π", -Math.PI, "-3.14159265"],
    ["123456789", 123_456_789, "123456789"],
    ["1e10", 1e10, "10000000000"],
    ["12345678901", 12_345_678_901, "12345678901"],
    ["99999999999", 99_999_999_999, "99999999999"],
    ["-12345678901 (12 characters fixed)", -12_345_678_901, "-1.2346E+10"],
    ["123456789012 (the spec's example)", 123_456_789_012, "1.23457E+11"],
    ["1e11", 1e11, "1E+11"],
    ["1.5e15", 1.5e15, "1.5E+15"],
    ["1e100", 1e100, "1E+100"],
    ["0.0000000001234 (the spec's example)", 0.000_000_000_123_4, "1.234E-10"],
    ["1e-9", 1e-9, "0.000000001"],
    ["1e-10", 1e-10, "1E-10"],
    ["-1e-10", -1e-10, "-1E-10"],
    ["1.5e-7", 1.5e-7, "0.00000015"],
    ["0.000123456789", 0.000_123_456_789, "0.000123457"],
    ["0.0000123456789012345", 0.000_012_345_678_901_234_5, "0.000012346"],
    ["45366.041666666664", 45_366.041666666664, "45366.04167"],
    ["12345678.12345", 12_345_678.12345, "12345678.12"],
    ["1234567890.12", 1_234_567_890.12, "1234567890"],
    ["-123456789.5", -123_456_789.5, "-123456790"],
    ["9999999999.99", 9_999_999_999.99, "10000000000"],
  ])("%s → %j", (_label, value, text) => {
    expect(generalText(value)).toBe(text);
  });

  it("renders every number in at most 11 characters, keeping its sign", () => {
    for (const n of SAMPLE) {
      const text = generalText(n);
      expect(text.length, `${n} → ${text}`).toBeLessThanOrEqual(WIDTH);
      expect(text.startsWith("-"), `${n} → ${text}`).toBe(n < 0);
    }
  });

  it("is scientific outside 1e-9 ≤ |n| < 1e11, the mantissa keeping as many digits as fit", () => {
    for (const n of SAMPLE) {
      const magnitude = Math.abs(n);
      if (magnitude >= 1e-9 && magnitude < 1e11) {
        continue;
      }
      const text = generalText(n);
      expect(text, String(n)).toMatch(SCIENTIFIC);
      expect(Math.abs(Number(text) - n), `${n} → ${text}`).toBeLessThanOrEqual(maxRoundingError(text) + magnitude * 1e-12);
      expect(Math.abs(Number(text) - n) / magnitude, `${n} → ${text}`).toBeLessThanOrEqual(1e-4);
    }
  });

  it("writes fixed notation only inside the range: n rounded to as many decimals as fit, no trailing zeros or point", () => {
    for (const n of SAMPLE) {
      const text = generalText(n);
      if (text.includes("E")) {
        continue;
      }
      const magnitude = Math.abs(n);
      expect(magnitude >= 1e-9 && magnitude < 1e11, `${n} → ${text}`).toBe(true);
      expect(text, String(n)).toMatch(FIXED);
      expect(Math.abs(Number(text) - n), `${n} → ${text}`).toBeLessThanOrEqual(maxRoundingError(text) + magnitude * 1e-12);
      // Below 1e-4 the 11 characters hold fewer than five significant digits
      // by design ("0.000000001" for 1.234e-9), so only the bound above holds.
      if (magnitude >= 1e-4) {
        expect(Math.abs(Number(text) - n) / magnitude, `${n} → ${text}`).toBeLessThanOrEqual(1e-4);
      }
    }
  });

  it("fits the extremes of the double range", () => {
    for (const n of [Number.MAX_VALUE, -Number.MAX_VALUE, Number.MIN_VALUE, -Number.MIN_VALUE, 1e-300, -1e300]) {
      const text = generalText(n);
      expect(text.length, `${n} → ${text}`).toBeLessThanOrEqual(WIDTH);
      expect(text, String(n)).toMatch(SCIENTIFIC);
    }
  });
});

// "If 1e-9 ≤ |n| < 1e11, the fixed-notation text of n rounded to at most 15
// significant digits, with trailing zeros and a trailing point removed; if
// that is longer than 11 characters (sign and point included), round off
// decimals until it fits."
describe("generalText — rounding off decimals until the text fits", () => {
  it.each<[value: number, text: string]>([
    // "12345678901.7" is 13 characters; with its decimal rounded off it is
    // the 11-character "12345678902". (No ties here, so no rounding-mode
    // question either.)
    [12_345_678_901.7, "12345678902"],
    [12_345_678_901.25, "12345678901"],
    // "-1234567890.75" → "-1234567891": 11 characters, sign included.
    [-1_234_567_890.75, "-1234567891"],
  ])("rounds off every decimal when only the integer part fits: %s → %j", (value, text) => {
    expect(generalText(value)).toBe(text);
  });

  it.each<[value: number, text: string]>([
    // The 15-significant-digit text is "0.1234567895" (12 characters); one
    // decimal rounded off is "0.12345679". (The double itself is
    // 0.12345678949999999707…, which rounds the other way.) The digit before
    // the 5 is odd, so half-up and half-even rounding agree.
    [0.123_456_789_5, "0.12345679"],
    // "1.0000000015" → "1.000000002" (the double is 1.00000000149999990207…).
    [1.000_000_001_5, "1.000000002"],
  ])("rounds off the decimals of the 15-significant-digit text: %s → %j", (value, text) => {
    expect(generalText(value)).toBe(text);
  });
});
