// Editing-surface geometry (SPEC-027): the page/surface size and margin
// types, the standard paper catalogue, orientation and margin validation.
// Product-neutral — a Slides product sizes its slides (16:9, 4:3, custom)
// with the same types and matching/validation logic, just a different
// catalogue of named sizes. The document model always stores geometry in
// centimetres; display-unit conversion lives in core/units.ts.

/** Page dimensions in centimetres. */
export interface PageSize {
  widthCm: number;
  heightCm: number;
}

/** Page margins in centimetres. */
export interface Margins {
  topCm: number;
  rightCm: number;
  bottomCm: number;
  leftCm: number;
}

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
 * plain string (espree bars the annotation in `.vue` scripts). */
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

// --- slide sizes (the Slides product's catalogue) ------------------------

/** The named slide sizes the deck-setup dialog offers (SPEC-090), plus the
 * custom escape hatch — the presentation-side counterpart of `PaperSizeName`.
 * `matchSlideSize` compares against `SLIDE_SIZES` in either orientation. */
export type SlideSizeName = "16:9" | "4:3" | "16:10" | "A4" | "Letter" | "Custom";

/** The standard slide sizes, LANDSCAPE-first (wider side first) — the reverse
 * of `PAPER_SIZES`, because slides default to landscape. The two paper entries
 * reuse `PAPER_SIZES`, swapped to landscape, rather than restating their
 * dimensions. Values in centimetres:
 *   16:9  → 12192000 × 6858000 EMU, 4:3 → 9144000 × 6858000 EMU. */
export const SLIDE_SIZES: Record<Exclude<SlideSizeName, "Custom">, PageSize> = {
  "16:9": { widthCm: 33.867, heightCm: 19.05 },
  "4:3": { widthCm: 25.4, heightCm: 19.05 },
  "16:10": { widthCm: 25.4, heightCm: 15.875 },
  A4: swapOrientation(PAPER_SIZES.A4),
  Letter: swapOrientation(PAPER_SIZES.Letter),
};

/** The catalogue order the preset selector lists (Custom last). */
export const SLIDE_SIZE_NAMES: SlideSizeName[] = ["16:9", "4:3", "16:10", "A4", "Letter", "Custom"];

/** The smallest / largest slide dimension the dialog accepts: 1 in and 56 in,
 * PowerPoint's own bounds. */
export const MIN_SLIDE_DIMENSION_CM = 2.54;
export const MAX_SLIDE_DIMENSION_CM = 142.24;

/** Typed pass-through for a slide-size name coming from a `.vue` template's
 * plain string (same espree constraint as `asPaperSizeName`). */
export function asSlideSizeName(value: string): SlideSizeName {
  return value as SlideSizeName;
}

/** Which standard slide size (in either orientation) matches `size` within
 * `MATCH_TOLERANCE_CM`, or `"Custom"` when none does — how the dialog
 * preselects the open deck's size (AC-4). */
export function matchSlideSize(size: PageSize): SlideSizeName {
  for (const name of SLIDE_SIZE_NAMES) {
    if (name === "Custom") {
      continue;
    }
    const candidate = SLIDE_SIZES[name];
    if (
      (approxEqual(size.widthCm, candidate.widthCm) && approxEqual(size.heightCm, candidate.heightCm)) ||
      (approxEqual(size.widthCm, candidate.heightCm) && approxEqual(size.heightCm, candidate.widthCm))
    ) {
      return name;
    }
  }
  return "Custom";
}

/** A named slide size's dimensions in the given orientation — the landscape
 * catalogue entry, swapped if portrait is requested. Custom has no catalogue
 * entry (its size comes from the width/height fields); callers only pass a
 * non-Custom name, but `name` stays the full `SlideSizeName` so a cast-back
 * `asSlideSizeName` value needs no second pass-through. */
export function slideSizeFor(name: SlideSizeName, orientation: Orientation): PageSize {
  const landscape = SLIDE_SIZES[name as Exclude<SlideSizeName, "Custom">] ?? SLIDE_SIZES["16:9"];
  return orientation === "portrait" ? swapOrientation(landscape) : { ...landscape };
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
