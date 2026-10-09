// The library card's preview of a workbook (SPEC-116): the active sheet's
// top-left corner as display strings, small enough to live in the store's
// shell record so the landing page renders a card without rejoining a
// single row. The store recomputes it whenever a sheet's rows reference
// changes (copy-on-write makes that the signal for "content changed").
//
// SPEC-118: each value is formatted by its cell's number format at full
// width in the locale it is given (formatValue, a fill as one character) —
// the app layer passes the user's preference, so a later locale change
// leaves a stored preview as it is until the next rebuild.
import { getCell, usedRange } from "./address";
import type { SheetPreview, SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import type { SpreadsheetLocale } from "./numberFormat/locale";
import { formatValue, plainText, workbookFormatContext } from "./numberFormat/render";
import { cellStyle, effectiveStyleId } from "./styles";

export const PREVIEW_ROWS = 12;
export const PREVIEW_COLS = 6;

/** Up to PREVIEW_ROWS × PREVIEW_COLS display strings from the active sheet's
 * top-left corner ("" for empty cells), trimmed to the sheet's used extent
 * inside that window; no cells for an empty or foreign sheet. */
export function buildPreview(
  model: Pick<SpreadsheetModel, "sheets" | "activeSheetIndex" | "styles" | "dateSystem" | "format" | "indexedColors">,
  locale: SpreadsheetLocale,
): SheetPreview {
  const sheet = model.sheets[model.activeSheetIndex] ?? model.sheets[0];
  if (!sheet) {
    return { sheetName: "", cells: [] };
  }
  const used = sheet.kind === "grid" ? usedRange(sheet) : null;
  if (!used) {
    return { sheetName: sheet.name, cells: [] };
  }
  const ctx = workbookFormatContext(model, locale);
  const rowCount = Math.min(PREVIEW_ROWS, used.bottom + 1);
  const colCount = Math.min(PREVIEW_COLS, used.right + 1);
  const cells: string[][] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const line: string[] = [];
    for (let col = 0; col < colCount; col += 1) {
      const value = getCell(sheet, row, col)?.value ?? null;
      const code = value === null ? undefined : formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row, col)), locale);
      line.push(plainText(formatValue(value, code, ctx)));
    }
    cells.push(line);
  }
  return { sheetName: sheet.name, cells };
}
