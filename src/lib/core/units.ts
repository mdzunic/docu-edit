// Unit conversions shared by every editing surface — pages today, slides
// tomorrow. One source for the fixed ratios that were previously spread
// across docx/units.ts (twips, EMU, half-points), pageSetup.ts (display
// units), images.ts (96 DPI px), pagination.ts (PX_PER_CM) and
// tableResize.ts (CM_PER_PX). The model's own units stay centimetres for
// layout and points for font sizes; everything here converts to/from that.

export const CM_PER_INCH = 2.54;
export const POINTS_PER_INCH = 72;

/** CSS reference pixels per inch — the fixed ratio browsers resolve
 * physical CSS units at, and the ratio SPEC-024 defines "Reset size" by. */
export const PX_PER_INCH = 96;

/** CSS reference pixels per centimetre — the same 96 dpi the canvas's
 * cm-sized page CSS resolves to, so engine px and rendered px agree. */
export const PX_PER_CM = PX_PER_INCH / CM_PER_INCH;

/** CSS's own fixed cm-to-px ratio (96 CSS px per inch): drag deltas convert
 * straight to centimetres without depending on measured geometry, which a
 * live width change can invalidate mid-drag. */
export const CM_PER_PX = CM_PER_INCH / PX_PER_INCH;

/** A natural pixel dimension to centimetres at 96 DPI. */
export function pxToCm(px: number): number {
  return (px / PX_PER_INCH) * CM_PER_INCH;
}

/** The inverse of pxToCm, for callers that need to compare against a pixel
 * measurement (e.g. a resize handle's on-screen delta). */
export function cmToPx(cm: number): number {
  return (cm / CM_PER_INCH) * PX_PER_INCH;
}

// --- display units (the Page setup dialog's selector) -----------------------

/** The unit a size/margin dialog displays values in (SPEC-027 AC-4, AC-6).
 * The model itself always stores cm — this only controls display/entry — so
 * no reader/writer maps it to anything in a `.docx` or `.odt` file. */
export type MeasurementUnit = "cm" | "mm" | "in" | "pt";

/** Typed pass-through for a measurement unit coming from a `.vue`
 * template's plain string (espree bars the annotation in `.vue` scripts). */
export function asMeasurementUnit(value: string): MeasurementUnit {
  return value as MeasurementUnit;
}

/** Centimetres per unit of measurement the selector offers: cm, mm, in, pt. */
const CM_PER_UNIT: Record<MeasurementUnit, number> = {
  cm: 1,
  mm: 0.1,
  in: CM_PER_INCH,
  pt: CM_PER_INCH / POINTS_PER_INCH,
};

/** A value in `unit` to centimetres — the model's own unit. */
export function toCm(value: number, unit: MeasurementUnit): number {
  return value * CM_PER_UNIT[unit];
}

/** Centimetres to a value in `unit`, for display in a dialog's fields. */
export function fromCm(cm: number, unit: MeasurementUnit): number {
  return cm / CM_PER_UNIT[unit];
}

/** Rounds a display value to a sensible number of decimals per unit: whole
 * points, hundredths of an inch/centimetre, tenths of a millimetre — fine
 * enough to look exact, coarse enough not to show float noise. */
export function roundForDisplay(value: number, unit: MeasurementUnit): number {
  const decimals = unit === "pt" ? 0 : unit === "mm" ? 1 : 2;
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

// --- OOXML integer units ----------------------------------------------------
// OOXML stores layout in fixed integer units. The conversions are shared
// with any OOXML-family package (.docx today, .pptx for a Slides product).

const TWIPS_PER_INCH = 1440;

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

// Inverse conversions for the writers (SPEC-006). Each rounds to the nearest
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

/** Eighths of a point (OOXML's unit for border width, `w:sz` on
 * `w:tcBorders`'s children) to points. */
export function eighthPointsToPt(eighths: number): number {
  return eighths / 8;
}

/** Points to eighths of a point (OOXML's unit for border width). */
export function ptToEighthPoints(pt: number): number {
  return Math.round(pt * 8);
}

// English Metric Units: OOXML's unit for DrawingML sizes (`wp:extent`'s
// cx/cy, SPEC-024), 914400 per inch — the same fixed conversion twips use
// (1440 per inch), just a different denominator.
const EMU_PER_INCH = 914400;

/** Centimetres to EMU, rounded to the nearest whole EMU. */
export function cmToEmu(cm: number): number {
  return Math.round((cm / CM_PER_INCH) * EMU_PER_INCH);
}

/** EMU to centimetres. */
export function emuToCm(emu: number): number {
  return (emu / EMU_PER_INCH) * CM_PER_INCH;
}
