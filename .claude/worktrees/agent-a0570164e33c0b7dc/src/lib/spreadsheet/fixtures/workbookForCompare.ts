// The facts SPEC-121's AC-3 compares when a written workbook is read back
// (`workbookForCompare` deep equality, "ids masked — compare resolved facts,
// not minted values"): per sheet, in order, its name, visibility, kind, tab
// colour resolved to hex (core/theme.ts's resolveColor against the model's
// theme, else DEFAULT_THEME — so a theme reference and the literal a format
// writes for it compare equal), default column width rounded to whole px and
// default row height to two decimals; the ACTIVE sheet's name (not its
// index-dependent id); the date system. Sheet ids, part names, retained
// source and every other minted value are left out.
import { DEFAULT_THEME, resolveColor } from "../../core/theme";
import type { Sheet, SpreadsheetModel } from "../model";

export interface ComparableSheet {
  name: string;
  visibility: Sheet["visibility"];
  kind: Sheet["kind"];
  tabColor: string | null;
  defaultColWidthPx: number;
  defaultRowHeightPt: number;
}

export interface ComparableWorkbook {
  sheets: ComparableSheet[];
  activeSheet: string | null;
  dateSystem: SpreadsheetModel["dateSystem"];
}

export function workbookForCompare(model: SpreadsheetModel): ComparableWorkbook {
  const theme = model.theme ?? DEFAULT_THEME;
  return {
    sheets: model.sheets.map((sheet) => ({
      name: sheet.name,
      visibility: sheet.visibility,
      kind: sheet.kind,
      tabColor: sheet.tabColor === undefined ? null : resolveColor(sheet.tabColor, theme).toLowerCase(),
      defaultColWidthPx: Math.round(sheet.defaultColWidthPx),
      defaultRowHeightPt: Math.round(sheet.defaultRowHeightPt * 100) / 100,
    })),
    activeSheet: model.sheets[model.activeSheetIndex]?.name ?? null,
    dateSystem: model.dateSystem,
  };
}
