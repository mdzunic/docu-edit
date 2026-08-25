// The font-family catalogue and font-size rules behind SPEC-013's toolbar
// controls. Framework-free (no Vue imports) per the src/lib convention.
// The standard catalogue lives in core/fonts.ts (product-neutral).
import { isBreak, isChart, isImage, isShape, isToc,
  isTof, type DocumentModel } from "./documentModel";

/** The families the open document actually uses, for the "Fonts in this
 * document" group (SPEC-013 AC): the document default plus every family a
 * run names, whether or not it is installed locally — the list comes from
 * the model, never from local font detection. Alphabetical, de-duplicated
 * case-insensitively (first spelling wins). */
export function documentFontFamilies(model: DocumentModel): string[] {
  const families = new Map<string, string>();
  const add = (family: string) => {
    const key = family.toLowerCase();
    if (!families.has(key)) {
      families.set(key, family);
    }
  };
  add(model.defaultStyles.fontFamily);
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (isImage(block) || isShape(block) || isChart(block) || isBreak(block) || isToc(block) || isTof(block)) {
        continue;
      }
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : [block];
      for (const paragraph of paragraphs) {
        for (const run of paragraph.runs) {
          if (run.formatting.fontFamily) {
            add(run.formatting.fontFamily);
          }
        }
      }
    }
  }
  return Array.from(families.values()).sort((a, b) => a.localeCompare(b));
}

// Word's font-size limits: 1pt to 409pt in half-point steps (SPEC-013 AC).
export const MIN_FONT_SIZE_PT = 1;
export const MAX_FONT_SIZE_PT = 409;

/** The common point sizes offered in the font-size control's dropdown
 * (SPEC-061 AC-21). The field stays free-typed for anything off this list;
 * these are just the quick picks a word processor conventionally shows. */
export const COMMON_FONT_SIZES = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72,
];

/** The range message for a given min/max (SPEC-081 reuses the control with a
 * narrower 6–96 presentation range, so the tooltip must name whichever limits
 * are in force, not always Word's 1–409). */
export function fontSizeRangeMessage(min = MIN_FONT_SIZE_PT, max = MAX_FONT_SIZE_PT): string {
  return `Font size must be between ${min} and ${max} pt`;
}

/** Shown when out-of-range input is clamped (SPEC-013 AC). */
export const FONT_SIZE_RANGE_MESSAGE = fontSizeRangeMessage();

export interface FontSizeParse {
  /** The accepted size, snapped to the 0.5 pt grid and clamped into range;
   * undefined when the input is not a number at all. */
  value?: number;
  /** True when the typed value was outside 1..409 and had to be clamped —
   * the control shows the range tooltip (SPEC-013 AC). */
  clamped: boolean;
}

/** Parses the size control's text input: accepts "11", "11.5", "11,5"
 * (decimal comma), trims whitespace, snaps to the 0.5 pt grid and clamps
 * into min..max (defaults to Word's 1..409; the presentation toolbar passes
 * its own 6..96). */
export function parseFontSizeInput(
  input: string,
  min = MIN_FONT_SIZE_PT,
  max = MAX_FONT_SIZE_PT,
): FontSizeParse {
  const raw = Number(input.trim().replace(",", "."));
  if (input.trim().length === 0 || Number.isNaN(raw)) {
    return { clamped: false };
  }
  const snapped = Math.round(raw * 2) / 2;
  const value = Math.min(Math.max(snapped, min), max);
  return { value, clamped: value !== snapped };
}

/** One +/− button step: the nearest 0.5 pt grid point strictly above/below
 * the current size (an off-grid 11.3 steps to 11.5 up, 11 down), clamped
 * into range. */
export function stepFontSize(
  current: number,
  direction: 1 | -1,
  min = MIN_FONT_SIZE_PT,
  max = MAX_FONT_SIZE_PT,
): number {
  const scaled = current * 2;
  const stepped = (direction === 1 ? Math.floor(scaled + 1) : Math.ceil(scaled - 1)) / 2;
  return Math.min(Math.max(stepped, min), max);
}
