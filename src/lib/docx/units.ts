// Unit conversions between OOXML's native units and the document model's
// units (centimetres for layout, points for font sizes).

const TWIPS_PER_INCH = 1440;
const CM_PER_INCH = 2.54;

/** Twentieths of a point (OOXML's unit for margins, page size and indents)
 * to centimetres. */
export function twipsToCm(twips: number): number {
  return (twips / TWIPS_PER_INCH) * CM_PER_INCH;
}

/** Half-points (OOXML's unit for font size, `w:sz`) to points. */
export function halfPointsToPt(halfPoints: number): number {
  return halfPoints / 2;
}

/** Twentieths of a point (OOXML's unit for spacing before/after and exact
 * line spacing) to points. */
export function twentiethsOfPointToPt(value: number): number {
  return value / 20;
}

/** Normalises an OOXML colour value ("FF0000", "auto") to a CSS colour, or
 * undefined for "auto"/empty (meaning "inherit the default"). */
export function ooxmlColorToCss(value: string | null): string | undefined {
  if (!value || value.toLowerCase() === "auto") {
    return undefined;
  }
  return `#${value.toLowerCase()}`;
}

// Inverse conversions for the writer (SPEC-006). Each rounds to the nearest
// whole OOXML unit: the model's centimetre/point floats came from integer
// twips/half-points in the first place, so rounding restores the original
// integer exactly and keeps open -> save -> re-open byte-stable.

/** Centimetres to twentieths of a point (OOXML's unit for margins, page
 * size and indents), rounded to the nearest whole twip. */
export function cmToTwips(cm: number): number {
  return Math.round((cm / CM_PER_INCH) * TWIPS_PER_INCH);
}

/** Points to half-points (OOXML's unit for font size, `w:sz`). */
export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

/** Points to twentieths of a point (OOXML's unit for spacing before/after
 * and exact line spacing). */
export function ptToTwentiethsOfPoint(pt: number): number {
  return Math.round(pt * 20);
}

/** A CSS `#rrggbb` colour to OOXML's bare uppercase hex ("FF0000"), or
 * undefined for anything else — the model only ever holds six-digit hex
 * colours (that's all `ooxmlColorToCss` produces), so any other shape has
 * no faithful OOXML encoding and is omitted rather than written broken. */
export function cssColorToOoxml(value: string | undefined): string | undefined {
  if (!value || !/^#[0-9a-f]{6}$/i.test(value)) {
    return undefined;
  }
  return value.slice(1).toUpperCase();
}
