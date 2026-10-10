// Row autofit (SPEC-122 "Row autofit") — THE normative source for the height
// a row's content needs and for the automatic refit every commit, Fill and
// Clear contents runs inside its undo step. SPEC-124's AutoFit commands
// reuse autofitRowHeightPt (and add autofitColumnWidthPx beside it).
//
// - A row is as tall as its tallest stored cell. A cell is one line of its
//   font per explicit "\n" line (unwrapped text never wraps otherwise);
//   wrapped text (alignment.wrap) wraps greedily at word boundaries — a word
//   wider than the line breaks between characters — at the column width
//   (the covering span's widthPx, else the sheet default) less
//   WRAP_PADDING_PX; rotated text (1–180) takes the height of its lines'
//   bounding box rotated by the angle (91–180: clockwise by r − 90);
//   stacked text (255) is its longest line's characters × the line height.
//   A rich-text run in a larger font raises the cell's line height. Cells in
//   a merged range are skipped (Excel's rule: autofit ignores merges).
// - The text measured is what the cell displays: SPEC-118's formatValue
//   through the effective style's code (formatCodeOf), as plain text — a
//   formula cell shows its value. The font is the effective style's, its
//   family resolved through core/theme.ts resolveFont against model.theme
//   (DEFAULT_THEME when absent; the theme's minor font when the style names
//   none), 11 pt when no size is set.
// - px → pt is × 0.75, rounded UP to a 0.75 pt step (one whole px), never
//   below the sheet's defaultRowHeightPt.
// - refitRows is the automatic refit: for the stored rows without
//   `customHeight` in `ranges` that hold a stored non-empty cell there — or
//   carry a height of their own, so a row whose content was cleared shrinks
//   back — up to AUTOFIT_MAX_ROWS_PER_COMMAND of them in row order, it sets
//   `heightPt` copy-on-write (absent when it equals the default; a changed
//   height drops the row's ods `sourceStyleName`, the drop rule) and never
//   `customHeight`. A row whose height does not change keeps its identity; a
//   row left with no cells and no properties is removed.
// - Measuring is injected (CellTextMeasurer): the view's canvas measurer
//   (src/views/spreadsheet/cellMeasurer.ts), or fixedMeasurer in tests.
import { DEFAULT_THEME, resolveFont } from "../core/theme";
import { findRow, type GridRange } from "./address";
import { replaceRows } from "./cells";
import { buildMergeIndex, mergeAt, type MergeIndex } from "./merges";
import type { Cell, CellFont, Row, Sheet, SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { formatValue, plainText, type FormatContext } from "./numberFormat/render";
import { cellStyle, columnSpanAt, effectiveStyleId } from "./styles";

export interface MeasuredFont {
  family: string;
  sizePt: number;
  bold: boolean;
  italic: boolean;
}

export interface CellTextMeasurer {
  widthPx(text: string, font: MeasuredFont): number;
  lineHeightPx(font: MeasuredFont): number;
}

/** The most rows one command refits. */
export const AUTOFIT_MAX_ROWS_PER_COMMAND = 10_000;

/** Wrapped text wraps at the column width less this padding. */
export const WRAP_PADDING_PX = 5;

/** Points per CSS pixel. */
const PT_PER_PX = 0.75;

/** A measurer where every character is `charPx` wide and every line
 * `linePx` tall, whatever the font (tests). */
export function fixedMeasurer(charPx: number, linePx: number): CellTextMeasurer {
  return {
    widthPx: (text) => [...text].length * charPx,
    lineHeightPx: () => linePx,
  };
}

/** A cell font as the measurer takes it (see the module header). */
export function measuredFont(model: Pick<SpreadsheetModel, "theme">, font: CellFont | undefined): MeasuredFont {
  const theme = model.theme ?? DEFAULT_THEME;
  return {
    family: font?.family !== undefined ? resolveFont(font.family, theme) : theme.fonts.minor,
    sizePt: font?.sizePt ?? 11,
    bold: font?.bold === true,
    italic: font?.italic === true,
  };
}

/** How many lines greedy word-wrapping makes of one paragraph. */
function wrappedLineCount(paragraph: string, availablePx: number, font: MeasuredFont, measurer: CellTextMeasurer): number {
  let count = 0;
  let current: string | null = null;
  for (const word of paragraph.split(" ")) {
    const candidate: string = current === null ? word : `${current} ${word}`;
    if (measurer.widthPx(candidate, font) <= availablePx) {
      current = candidate;
      continue;
    }
    if (current !== null) {
      count += 1;
    }
    // The word starts a line; while it is wider than one, it breaks after
    // the longest prefix that fits (at least one character).
    let rest = [...word];
    while (rest.length > 1 && measurer.widthPx(rest.join(""), font) > availablePx) {
      let lo = 1;
      let hi = rest.length - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >>> 1;
        if (measurer.widthPx(rest.slice(0, mid).join(""), font) <= availablePx) {
          lo = mid;
        } else {
          hi = mid - 1;
        }
      }
      count += 1;
      rest = rest.slice(lo);
    }
    current = rest.join("");
  }
  return count + 1;
}

/** The text a cell displays (formula cells: their value). */
function displayText(cell: Cell, code: string | undefined, ctx: FormatContext): string {
  return plainText(formatValue(cell.value, code, { ...ctx, widthPx: undefined }));
}

/** The px height one cell's content needs. */
function cellHeightPx(model: SpreadsheetModel, sheet: Sheet, row: number, cell: Cell, measurer: CellTextMeasurer, ctx: FormatContext): number {
  const style = cellStyle(model, effectiveStyleId(sheet, row, cell.col));
  const font = measuredFont(model, style.font);
  let lineHeight = measurer.lineHeightPx(font);
  for (const run of cell.runs ?? []) {
    if (run.font) {
      lineHeight = Math.max(lineHeight, measurer.lineHeightPx(measuredFont(model, { ...style.font, ...run.font })));
    }
  }
  const lines = displayText(cell, formatCodeOf(style, ctx.locale), ctx).split("\n");
  const rotation = style.alignment?.textRotation ?? 0;
  if (rotation === 255) {
    const characters = Math.max(1, ...lines.map((line) => [...line].length));
    return characters * lineHeight;
  }
  if (rotation >= 1 && rotation <= 180) {
    const width = Math.max(0, ...lines.map((line) => measurer.widthPx(line, font)));
    const height = lines.length * lineHeight;
    const angle = ((rotation <= 90 ? rotation : rotation - 90) * Math.PI) / 180;
    return Math.abs(width * Math.sin(angle)) + Math.abs(height * Math.cos(angle));
  }
  if (style.alignment?.wrap) {
    const columnPx = columnSpanAt(sheet.columns, cell.col)?.widthPx ?? sheet.defaultColWidthPx;
    const available = Math.max(1, columnPx - WRAP_PADDING_PX);
    let count = 0;
    for (const line of lines) {
      count += wrappedLineCount(line, available, font, measurer);
    }
    return count * lineHeight;
  }
  return lines.length * lineHeight;
}

/** Each merges array's index, built once (arrays are replaced, never
 * mutated). */
const mergeIndexes = new WeakMap<readonly GridRange[], MergeIndex>();

function mergeIndexOf(sheet: Sheet): MergeIndex {
  let index = mergeIndexes.get(sheet.merges);
  if (!index) {
    index = buildMergeIndex(sheet.merges);
    mergeIndexes.set(sheet.merges, index);
  }
  return index;
}

/** px → pt, rounded up to a 0.75 pt step (one whole px). */
function pxToRowPt(px: number): number {
  return Math.ceil(px - 1e-7) * PT_PER_PX;
}

/** The height (pt) row `row`'s content needs (see the module header). */
export function autofitRowHeightPt(model: SpreadsheetModel, sheet: Sheet, row: number, measurer: CellTextMeasurer, ctx: FormatContext): number {
  const stored = findRow(sheet, row);
  let tallest = 0;
  if (stored) {
    const merges = mergeIndexOf(sheet);
    for (const cell of stored.cells) {
      if (merges.ranges.length > 0 && mergeAt(merges, row, cell.col)) {
        continue;
      }
      tallest = Math.max(tallest, cellHeightPx(model, sheet, row, cell, measurer, ctx));
    }
  }
  return Math.max(pxToRowPt(tallest), sheet.defaultRowHeightPt);
}

/** Whether `row` holds a stored non-empty cell inside one of `ranges`. */
function holdsContentIn(row: Row, ranges: readonly GridRange[]): boolean {
  return ranges.some((range) => {
    if (row.index < range.top || row.index > range.bottom) {
      return false;
    }
    return row.cells.some((cell) => cell.col >= range.left && cell.col <= range.right && (cell.value !== null || cell.formula !== undefined));
  });
}

/** The automatic refit of the rows `ranges` touched (see the module
 * header). */
export function refitRows(model: SpreadsheetModel, sheetId: string, ranges: GridRange[], measurer: CellTextMeasurer, ctx: FormatContext): void {
  const sheetIndex = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[sheetIndex];
  if (!sheet || ranges.length === 0) {
    return;
  }
  const top = Math.min(...ranges.map((range) => range.top));
  const bottom = Math.max(...ranges.map((range) => range.bottom));
  const replaced = new Map<number, Row | undefined>();
  let refitted = 0;
  const rows = sheet.rows;
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (rows[mid]!.index < top) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  for (let r = lo; r < rows.length; r += 1) {
    const row = rows[r]!;
    if (row.index > bottom || refitted >= AUTOFIT_MAX_ROWS_PER_COMMAND) {
      break;
    }
    if (row.customHeight || !ranges.some((range) => row.index >= range.top && row.index <= range.bottom)) {
      continue;
    }
    if (!holdsContentIn(row, ranges) && row.heightPt === undefined) {
      continue;
    }
    refitted += 1;
    const height = autofitRowHeightPt(model, sheet, row.index, measurer, ctx);
    const heightPt = height === sheet.defaultRowHeightPt ? undefined : height;
    if (heightPt === row.heightPt) {
      continue;
    }
    const next: Row = { ...row };
    if (heightPt === undefined) {
      delete next.heightPt;
    } else {
      next.heightPt = heightPt;
    }
    delete next.sourceStyleName;
    const empty = next.cells.length === 0 && Object.entries(next).every(([key, value]) => key === "index" || key === "cells" || value === undefined);
    replaced.set(row.index, empty ? undefined : next);
  }
  if (replaced.size > 0) {
    replaceRows(model, sheetIndex, replaced);
  }
}
