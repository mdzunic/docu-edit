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
