// Unit conversions between ODF's native value formats and the document
// model's units (centimetres for layout, points for font sizes). Unlike
// OOXML's fixed integer units, ODF writes measured values as strings with an
// explicit unit suffix ("2.54cm", "1in", "12pt") and expresses some
// properties as percentages of an inherited value ("130%").

import { CM_PER_INCH, POINTS_PER_INCH } from "../core/units";
import { odfLengthToPt, odfPercentage } from "../core/odfXml";

// The measured-value PARSERS (odfLengthToCm/odfLengthToPt/odfPercentage/
// odfColorToCss) moved to core/odfXml.ts (SPEC-079) so the product-neutral
// core and the `.odp` reader share them; they are re-exported here so every
// existing `odf/*` importer keeps its import path. The WRITER direction (the
// byte-exact round-trip searches below) stays odf-local.
export { odfLengthToCm, odfLengthToPt, odfPercentage, odfColorToCss } from "../core/odfXml";
// `cmToOdfLength` moved to core/odfXml.ts (SPEC-092) so the hoisted chart-part
// codec can emit `svg:width` without importing `../odf/*`; re-exported here so
// every existing `odf/*` importer keeps its path.
export { cmToOdfLength } from "../core/odfXml";

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
