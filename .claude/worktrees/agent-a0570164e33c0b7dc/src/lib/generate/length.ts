// Length helper inputs (SPEC-045 AC-25, AC-26): every length the builder API
// or the JSON document-description format accepts is one of four unit keys,
// converted to an integer Emu (English Metric Units, 914400/inch — see
// specs/002-create-blank-document.md's "Canonical document model" for the
// unit) on entry. `cm`/`pt`/`in` are physical measurements rounded to the
// nearest whole Emu, exactly like docx/units.ts's cmToEmu; `emu` is already
// the target unit, so a fractional or negative value there is a caller
// error, not something to silently round.
import { DocumentGenerationError } from "./errors";

export interface LengthCm {
  cm: number;
}
export interface LengthPt {
  pt: number;
}
export interface LengthIn {
  in: number;
}
export interface LengthEmu {
  emu: number;
}

/** A length in one of four units — the only shape the generation API ever
 * accepts for a measurement, never a bare number (which unit it would be is
 * ambiguous). */
export type LengthInput = LengthCm | LengthPt | LengthIn | LengthEmu;

const EMU_PER_CM = 360000;
const EMU_PER_POINT = 12700;
const EMU_PER_INCH = 914400;

const LENGTH_UNITS = ["cm", "pt", "in", "emu"] as const;
type LengthUnit = (typeof LENGTH_UNITS)[number];

function isLengthUnit(value: string): value is LengthUnit {
  return (LENGTH_UNITS as readonly string[]).includes(value);
}

/** Validates and converts a length input to an integer Emu (AC-25). Throws
 * `DocumentGenerationError` naming `path` for anything but a plain object
 * with exactly one of `cm`/`pt`/`in`/`emu` set to a non-negative finite
 * number — a non-integer or negative value is always rejected (AC-26), and
 * `emu` additionally requires an integer since no conversion happens for it. */
export function lengthToEmu(value: unknown, path: string): number {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new DocumentGenerationError(
      `${path}: expected a length ({ cm }, { pt }, { in } or { emu })`,
    );
  }
  const keys = Object.keys(value as Record<string, unknown>);
  if (keys.length !== 1 || !isLengthUnit(keys[0])) {
    throw new DocumentGenerationError(
      `${path}: expected exactly one of "cm", "pt", "in" or "emu"`,
    );
  }
  const unit = keys[0];
  const raw = (value as Record<string, unknown>)[unit];
  if (typeof raw !== "number" || !Number.isFinite(raw)) {
    throw new DocumentGenerationError(`${path}.${unit}: expected a finite number`);
  }
  if (raw < 0) {
    throw new DocumentGenerationError(`${path}.${unit}: expected a non-negative number`);
  }
  if (unit === "emu") {
    if (!Number.isInteger(raw)) {
      throw new DocumentGenerationError(`${path}.emu: expected an integer`);
    }
    return raw;
  }
  if (unit === "cm") {
    return Math.round(raw * EMU_PER_CM);
  }
  if (unit === "pt") {
    return Math.round(raw * EMU_PER_POINT);
  }
  return Math.round(raw * EMU_PER_INCH);
}

/** Integer Emu to centimetres — the unit `ParagraphFormatting`/`PageSize`/
 * `Table.columnWidthsCm`/`Image` width/height fields store (SPEC-002's
 * canonical model measures layout in cm, not Emu; Emu is this API's own
 * input boundary, converted immediately on entry per AC-25). */
export function emuToCm(emu: number): number {
  return emu / EMU_PER_CM;
}

/** Integer Emu to points — the unit `RunFormatting.fontSizePt`/paragraph
 * spacing fields store. */
export function emuToPt(emu: number): number {
  return emu / EMU_PER_POINT;
}
