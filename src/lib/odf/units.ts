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
