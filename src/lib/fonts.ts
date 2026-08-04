// The font-family catalogue and font-size rules behind SPEC-013's toolbar
// controls. Framework-free (no Vue imports) per the src/lib convention.
import type { DocumentModel } from "./documentModel";

/** The families the font dropdown always offers (SPEC-013 AC): the classic
 * web-safe set plus the common document fonts — Word's modern defaults
 * (Calibri/Cambria) and their metric-compatible Liberation/Carlito
 * equivalents. Alphabetical, so the list is stable and scannable. */
export const STANDARD_FONT_FAMILIES: string[] = [
  "Arial",
  "Calibri",
  "Cambria",
  "Carlito",
  "Courier New",
  "Garamond",
  "Georgia",
  "Helvetica",
  "Liberation Mono",
  "Liberation Sans",
  "Liberation Serif",
  "Tahoma",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

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
      for (const run of block.runs) {
        if (run.formatting.fontFamily) {
          add(run.formatting.fontFamily);
        }
      }
    }
  }
  return Array.from(families.values()).sort((a, b) => a.localeCompare(b));
}

// Word's font-size limits: 1pt to 409pt in half-point steps (SPEC-013 AC).
export const MIN_FONT_SIZE_PT = 1;
export const MAX_FONT_SIZE_PT = 409;
export const FONT_SIZE_STEP_PT = 0.5;

/** Shown when out-of-range input is clamped (SPEC-013 AC). */
export const FONT_SIZE_RANGE_MESSAGE =
  `Font size must be between ${MIN_FONT_SIZE_PT} and ${MAX_FONT_SIZE_PT} pt`;

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
 * into 1..409. */
export function parseFontSizeInput(input: string): FontSizeParse {
  const raw = Number(input.trim().replace(",", "."));
  if (input.trim().length === 0 || Number.isNaN(raw)) {
    return { clamped: false };
  }
  const snapped = Math.round(raw * 2) / 2;
  const value = Math.min(Math.max(snapped, MIN_FONT_SIZE_PT), MAX_FONT_SIZE_PT);
  return { value, clamped: value !== snapped };
}

/** One +/− button step: the nearest 0.5 pt grid point strictly above/below
 * the current size (an off-grid 11.3 steps to 11.5 up, 11 down), clamped
 * into range. */
export function stepFontSize(current: number, direction: 1 | -1): number {
  const scaled = current * 2;
  const stepped = (direction === 1 ? Math.floor(scaled + 1) : Math.ceil(scaled - 1)) / 2;
  return Math.min(Math.max(stepped, MIN_FONT_SIZE_PT), MAX_FONT_SIZE_PT);
}
