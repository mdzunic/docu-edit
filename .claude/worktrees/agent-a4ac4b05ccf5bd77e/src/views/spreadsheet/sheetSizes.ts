// Row heights and column widths as the size dialogs and the resize tooltip
// show them (SPEC-124 "Sizes"): rows in points (two decimals), columns in
// Excel characters through SPEC-117's MDW formula, each with its pixel
// equivalent at 100% zoom. Framework-free.
import { charsToXlsxWidth, pxToChars, roundRowHeightPt, xlsxWidthToChars, type Band, type BandAxis } from "../../lib/spreadsheet/bands";
import { ptToPx, xlsxWidthToPx } from "../../lib/spreadsheet/geometry";
import type { Sheet } from "../../lib/spreadsheet/model";
import { firstRowAtOrAfter } from "../../lib/spreadsheet/restyle";

/** A number as the dialogs show it: at most two decimals, no trailing 0s. */
export function sizeText(amount: number): string {
  return String(Number(amount.toFixed(2)));
}

/** The value a size dialog opens with: the bands' common height (pt) or
 * width (characters), "" when they differ. */
export function initialSize(sheet: Sheet, axis: BandAxis, bands: readonly Band[], mdw: number): string {
  const seen = new Set<number>();
  if (axis === "row") {
    for (const band of bands) {
      let stored = 0;
      for (let r = firstRowAtOrAfter(sheet.rows, band.start); r < sheet.rows.length && sheet.rows[r]!.index <= band.end; r += 1) {
        seen.add(sheet.rows[r]!.heightPt ?? sheet.defaultRowHeightPt);
        stored += 1;
      }
      if (stored < band.end - band.start + 1) {
        seen.add(sheet.defaultRowHeightPt);
      }
    }
    return seen.size === 1 ? sizeText([...seen][0]!) : "";
  }
  const defaultChars = pxToChars(sheet.defaultColWidthPx, mdw);
  for (const band of bands) {
    let col = band.start;
    for (const span of sheet.columns) {
      if (span.max < band.start || span.min > band.end) {
        continue;
      }
      if (span.min > col) {
        seen.add(defaultChars);
      }
      seen.add(span.sourceWidthChars !== undefined ? xlsxWidthToChars(span.sourceWidthChars, mdw) : pxToChars(span.widthPx ?? sheet.defaultColWidthPx, mdw));
      col = span.max + 1;
    }
    if (col <= band.end) {
      seen.add(defaultChars);
    }
  }
  return seen.size === 1 ? sizeText([...seen][0]!) : "";
}

/** The pixels a dialog amount stands for: points (rounded to 0.75 pt) for
 * rows, characters for columns. */
export function sizePixels(axis: BandAxis, amount: number, mdw: number): number {
  if (!(amount > 0)) {
    return 0;
  }
  return axis === "row" ? ptToPx(roundRowHeightPt(amount)) : xlsxWidthToPx(charsToXlsxWidth(amount, mdw), mdw);
}

/** The drag tooltip: "Width: 8.43 (64 pixels)" / "Height: 15.00 (20 pixels)". */
export function resizeTooltip(axis: BandAxis, px: number, mdw: number): string {
  const size = Math.max(0, Math.round(px));
  if (axis === "row") {
    return `Height: ${(size * 0.75).toFixed(2)} (${size} pixels)`;
  }
  return `Width: ${pxToChars(size, mdw).toFixed(2)} (${size} pixels)`;
}
