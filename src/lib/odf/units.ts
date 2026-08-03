// Unit conversions between ODF's native value formats and the document
// model's units (centimetres for layout, points for font sizes). Unlike
// OOXML's fixed integer units, ODF writes measured values as strings with an
// explicit unit suffix ("2.54cm", "1in", "12pt") and expresses some
// properties as percentages of an inherited value ("130%").

const CM_PER_INCH = 2.54;
const POINTS_PER_INCH = 72;

// Multipliers to centimetres for every unit ODF's positive/signed length
// datatypes allow (px is not a legal ODF length and is deliberately absent).
const CM_PER_UNIT: Record<string, number> = {
  cm: 1,
  mm: 0.1,
  in: CM_PER_INCH,
  pt: CM_PER_INCH / POINTS_PER_INCH,
  pc: (CM_PER_INCH / POINTS_PER_INCH) * 12,
};

const LENGTH_PATTERN = /^(-?\d+(?:\.\d+)?)(cm|mm|in|pt|pc)$/;

/** Parses an ODF length ("1.27cm", "-0.5in", "12pt") to centimetres, or
 * undefined for a missing/unitless/percentage value. */
export function odfLengthToCm(value: string | null | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = LENGTH_PATTERN.exec(value.trim());
  if (!match) {
    return undefined;
  }
  return Number(match[1]) * CM_PER_UNIT[match[2]];
}

/** Parses an ODF length to points (the model's unit for font sizes and
 * paragraph spacing), or undefined if it is not an absolute length. */
export function odfLengthToPt(value: string | null | undefined): number | undefined {
  const cm = odfLengthToCm(value);
  return cm === undefined ? undefined : (cm / CM_PER_INCH) * POINTS_PER_INCH;
}

const PERCENTAGE_PATTERN = /^(\d+(?:\.\d+)?)%$/;

/** Parses an ODF percentage ("130%") to its number (130), or undefined for
 * anything else. Some ODF properties (heading font sizes, line height) are
 * percentages of an inherited value rather than absolute lengths. */
export function odfPercentage(value: string | null | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = PERCENTAGE_PATTERN.exec(value.trim());
  if (!match) {
    return undefined;
  }
  return Number(match[1]);
}

/** Normalises an ODF colour ("#C9211E") to a CSS colour ("#c9211e"), or
 * undefined for "transparent"/keyword/empty values (meaning "inherit"). */
export function odfColorToCss(value: string | null | undefined): string | undefined {
  if (!value || !value.startsWith("#")) {
    return undefined;
  }
  return value.toLowerCase();
}

// --- writer direction ------------------------------------------------------
// The writer's requirement is stricter than "format the number": the save
// round-trip tests compare models with toEqual, so parsing the emitted
// string must reproduce the model's exact double. Formatting-then-parsing
// is not the identity in floating point (the pt parse path rounds three
// times), so each formatter searches for a string whose parse is bit-exact,
// preferring the shortest decimal.

/** `value.toFixed(digits)` with trailing fraction zeros trimmed. */
function trimmedFixed(value: number, digits: number): string {
  const text = value.toFixed(digits);
  return text.includes(".") ? text.replace(/0+$/, "").replace(/\.$/, "") : text;
}

const FLOAT_VIEW = new Float64Array(1);
const BITS_VIEW = new BigUint64Array(FLOAT_VIEW.buffer);

/** The double `steps` representable values away from `value` (positive
 * steps move toward +Infinity), saturating at zero rather than crossing it —
 * candidates for the neighbourhood searches below. */
function ulpStep(value: number, steps: number): number {
  FLOAT_VIEW[0] = value;
  const bits = BITS_VIEW[0];
  const negative = (bits & 0x8000000000000000n) !== 0n;
  let magnitude = bits & 0x7fffffffffffffffn;
  magnitude += negative ? BigInt(-steps) : BigInt(steps);
  if (magnitude < 0n) {
    magnitude = 0n;
  }
  BITS_VIEW[0] = negative ? magnitude | 0x8000000000000000n : magnitude;
  return FLOAT_VIEW[0];
}

/** Formats a centimetre value as an ODF length ("1.27cm") that
 * `odfLengthToCm` parses back to exactly `cm`. Exactness is guaranteed: the
 * cm parse path multiplies by 1, so the full-precision decimal always
 * round-trips — the search just prefers the shortest form. */
export function cmToOdfLength(cm: number): string {
  for (let digits = 0; digits <= 17; digits++) {
    const candidate = `${trimmedFixed(cm, digits)}cm`;
    if (odfLengthToCm(candidate) === cm) {
      return candidate;
    }
  }
  return `${cm}cm`;
}

/** Formats a point value as an ODF length ("12pt") that `odfLengthToPt`
 * parses back to exactly `pt`. Model pt values come from two places, both
 * covered: values parsed from a pt-denominated attribute are found by the
 * pt-unit search, and values parsed from a cm-denominated one (LibreOffice
 * writes paragraph spacing in cm) by the cm-unit search. Percentage-derived
 * font sizes may need a decimal one ulp off the value itself, hence the
 * neighbourhood walks. The final fallback — for doubles no ODF length
 * string can parse to, which therefore never came from a file — is at most
 * one ulp off. */
export function ptToOdfLength(pt: number): string {
  for (let digits = 0; digits <= 17; digits++) {
    const candidate = `${trimmedFixed(pt, digits)}pt`;
    if (odfLengthToPt(candidate) === pt) {
      return candidate;
    }
  }
  for (let steps = -6; steps <= 6; steps++) {
    const candidate = `${ulpStep(pt, steps)}pt`;
    if (odfLengthToPt(candidate) === pt) {
      return candidate;
    }
  }
  // One ulp of the cm-scale base moves the parsed pt result by less than an
  // ulp of `pt`, so this window spans the same neighbourhood as the pt walk.
  const cmBase = (pt / POINTS_PER_INCH) * CM_PER_INCH;
  for (let digits = 0; digits <= 17; digits++) {
    const candidate = `${trimmedFixed(cmBase, digits)}cm`;
    if (odfLengthToPt(candidate) === pt) {
      return candidate;
    }
  }
  for (let steps = -34; steps <= 34; steps++) {
    const candidate = `${ulpStep(cmBase, steps)}cm`;
    if (odfLengthToPt(candidate) === pt) {
      return candidate;
    }
  }
  return `${pt}pt`;
}

/** Formats a line-height multiplier (1.15) as an ODF percentage ("115%")
 * that `odfPercentage` divided by 100 parses back to exactly the
 * multiplier — the inverse of the reader's fo:line-height mapping. */
export function multiplierToOdfPercentage(value: number): string {
  const percent = value * 100;
  for (let digits = 0; digits <= 17; digits++) {
    const candidate = `${trimmedFixed(percent, digits)}%`;
    const parsed = odfPercentage(candidate);
    if (parsed !== undefined && parsed / 100 === value) {
      return candidate;
    }
  }
  return `${percent}%`;
}
