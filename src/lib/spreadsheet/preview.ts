// The library card's preview of a workbook (SPEC-116): the active sheet's
// top-left corner as General-format display strings, small enough to live in
// the store's shell record so the landing page renders a card without
// rejoining a single row. The store recomputes it whenever a sheet's rows
// reference changes (copy-on-write makes that the signal for "content
// changed").
import { getCell, usedRange } from "./address";
import { generalText } from "./general";
import type { SheetPreview, SpreadsheetModel } from "./model";

export const PREVIEW_ROWS = 12;
export const PREVIEW_COLS = 6;

/** Up to PREVIEW_ROWS × PREVIEW_COLS display strings from the active sheet's
 * top-left corner ("" for empty cells), trimmed to the sheet's used extent
 * inside that window; no cells for an empty or foreign sheet. */
export function buildPreview(model: Pick<SpreadsheetModel, "sheets" | "activeSheetIndex">): SheetPreview {
  const sheet = model.sheets[model.activeSheetIndex] ?? model.sheets[0];
  if (!sheet) {
    return { sheetName: "", cells: [] };
  }
  const used = sheet.kind === "grid" ? usedRange(sheet) : null;
  if (!used) {
    return { sheetName: sheet.name, cells: [] };
  }
  const rowCount = Math.min(PREVIEW_ROWS, used.bottom + 1);
  const colCount = Math.min(PREVIEW_COLS, used.right + 1);
  const cells: string[][] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const line: string[] = [];
    for (let col = 0; col < colCount; col += 1) {
      line.push(generalText(getCell(sheet, row, col)?.value ?? null));
    }
    cells.push(line);
  }
  return { sheetName: sheet.name, cells };
}
