// Paper sizes, orientation and unit conversions for the Page setup dialog
// (SPEC-027). Framework-free so the matching/validation logic is testable
// without mounting the dialog. The document model always stores page
// geometry in centimetres (documentModel.ts's PageSize/Margins) — this
// module only converts between that and what the dialog shows/accepts.
import type { Margins, MeasurementUnit, PageSize } from "./documentModel";

export type PaperSizeName = "A4" | "Letter" | "Legal" | "A3" | "A5" | "Tabloid" | "Custom";

/** The standard sizes the dialog offers (AC-1), portrait orientation
 * (narrower side first) — the canonical form `matchPaperSize` compares
 * against and `pageSizeFor` starts from. */
export const PAPER_SIZES: Record<Exclude<PaperSizeName, "Custom">, PageSize> = {
  A4: { widthCm: 21, heightCm: 29.7 },
  Letter: { widthCm: 21.59, heightCm: 27.94 },
  Legal: { widthCm: 21.59, heightCm: 35.56 },
  A3: { widthCm: 29.7, heightCm: 42 },
  A5: { widthCm: 14.8, heightCm: 21 },
  Tabloid: { widthCm: 27.94, heightCm: 43.18 },
};

export const PAPER_SIZE_NAMES: PaperSizeName[] = [
  "A4",
  "Letter",
  "Legal",
  "A3",
  "A5",
  "Tabloid",
  "Custom",
];

/** Typed pass-through for a paper size name coming from a `.vue` template's
 * plain string (same espree constraint as documentModel.ts's
 * asCharacterFormatKey). */
export function asPaperSizeName(value: string): PaperSizeName {
  return value as PaperSizeName;
}

// A source page size that came through the twip grid (.docx's pgSz) can be
// a couple hundredths of a centimetre off a standard size's exact value —
// well inside a fifth of a millimetre, which is the tolerance used here.
const MATCH_TOLERANCE_CM = 0.05;

function approxEqual(a: number, b: number): boolean {
  return Math.abs(a - b) < MATCH_TOLERANCE_CM;
}

/** Which standard paper size (in either orientation) matches `pageSize`
 * within twip-rounding tolerance, or `"Custom"` when it matches none (AC-8:
 * populating the dialog for a document whose size is a non-standard
 * custom size). */
export function matchPaperSize(pageSize: PageSize): PaperSizeName {
  for (const name of PAPER_SIZE_NAMES) {
    if (name === "Custom") {
      continue;
    }
    const size = PAPER_SIZES[name];
    if (
      (approxEqual(pageSize.widthCm, size.widthCm) && approxEqual(pageSize.heightCm, size.heightCm)) ||
      (approxEqual(pageSize.widthCm, size.heightCm) && approxEqual(pageSize.heightCm, size.widthCm))
    ) {
      return name;
    }
  }
  return "Custom";
}

export type Orientation = "portrait" | "landscape";

/** Typed pass-through for an orientation coming from a `.vue` template's
 * plain string (same espree constraint as asPaperSizeName). */
export function asOrientation(value: string): Orientation {
  return value as Orientation;
}

/** A page is landscape iff it is wider than it is tall — the same
 * convention docx/writer.ts and odf/writer.ts already use for `w:orient`/
 * `style:print-orientation`, so this always agrees with what gets saved. */
export function orientationOf(pageSize: PageSize): Orientation {
  return pageSize.widthCm > pageSize.heightCm ? "landscape" : "portrait";
}

/** Swaps width and height — the orientation toggle's effect (AC-2),
 * applied to whatever size is currently selected (a standard size or a
 * custom one), so it never needs to know which. */
export function swapOrientation(pageSize: PageSize): PageSize {
  return { widthCm: pageSize.heightCm, heightCm: pageSize.widthCm };
}

/** A named standard size's dimensions in the given orientation — the
 * portrait catalogue entry, swapped if landscape is requested. Callers only
 * ever pass a non-Custom name (Custom has no catalogue entry — its size
 * comes from the dialog's own width/height fields instead); `name` is
 * still typed as the full `PaperSizeName` so call sites fed a cast-back
 * `asPaperSizeName` value (from a `.vue` select's plain string) don't need
 * a second, narrower pass-through just for this one call. */
export function pageSizeFor(name: PaperSizeName, orientation: Orientation): PageSize {
  const portrait = PAPER_SIZES[name as Exclude<PaperSizeName, "Custom">] ?? PAPER_SIZES.A4;
  return orientation === "landscape" ? swapOrientation(portrait) : { ...portrait };
}

// --- units -------------------------------------------------------------

const CM_PER_INCH = 2.54;
const POINTS_PER_INCH = 72;

/** Centimetres per unit of measurement the dialog's selector offers
 * (AC-4): cm, mm, in and pt. */
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

/** Centimetres to a value in `unit`, for display in the dialog's fields. */
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

// --- validation ----------------------------------------------------------

/** The minimum width or height a text area may have (AC-9): margins that
 * would leave less than this are rejected. */
export const MIN_TEXT_AREA_CM = 1;

/** The exact message AC-9 requires for margins that leave too little room. */
export const MARGINS_TOO_LARGE_MESSAGE = "Margins are too large for this page size";

/** Whether `margins` would leave a text area (in either dimension) narrower
 * than `MIN_TEXT_AREA_CM` on `pageSize` (AC-9). */
export function textAreaTooNarrow(pageSize: PageSize, margins: Margins): boolean {
  const width = pageSize.widthCm - margins.leftCm - margins.rightCm;
  const height = pageSize.heightCm - margins.topCm - margins.bottomCm;
  return width < MIN_TEXT_AREA_CM || height < MIN_TEXT_AREA_CM;
}
