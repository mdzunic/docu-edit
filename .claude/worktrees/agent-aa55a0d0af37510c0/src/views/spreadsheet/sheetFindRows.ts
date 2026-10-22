// Find all's rows (SPEC-134): one per matching cell, in search order, up to
// FIND_ALL_LIMIT — its sheet's name, its address, its value as the cell shows
// it (the Values search text) and its formula ("=" + formula, "" for a
// constant) — with the cap message when there are more. Framework-free;
// useSheetFind builds them, SheetFindResults.vue lists them.
import { formatA1, getCell } from "../../lib/spreadsheet/address";
import { findAllCapMessage, FIND_ALL_LIMIT, searchText, type FindHit } from "../../lib/spreadsheet/find";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";

export interface FindResultRow {
  /** The hit's index (activating the row selects it). */
  index: number;
  sheetId: string;
  sheetName: string;
  ref: string;
  value: string;
  formula: string;
}

export interface FindResults {
  rows: FindResultRow[];
  total: number;
  /** "Showing the first 5,000 of N matches", or "" under the limit. */
  cap: string;
}

export function findResultRows(model: SpreadsheetModel, hits: readonly FindHit[], ctx: FormatContext): FindResults {
  const sheets = new Map(model.sheets.map((sheet) => [sheet.id, sheet]));
  const rows = hits.slice(0, FIND_ALL_LIMIT).map((hit, index): FindResultRow => {
    const sheet = sheets.get(hit.sheetId);
    const cell = sheet ? getCell(sheet, hit.row, hit.col) : undefined;
    return {
      index,
      sheetId: hit.sheetId,
      sheetName: sheet?.name ?? "",
      ref: formatA1(hit),
      value: sheet && cell ? (searchText(model, sheet, cell, "values", ctx, hit.row) ?? "") : "",
      formula: cell?.formula === undefined ? "" : `=${cell.formula}`,
    };
  });
  return { rows, total: hits.length, cap: hits.length > FIND_ALL_LIMIT ? findAllCapMessage(hits.length) : "" };
}

/** Narrows SheetFindResults' runtime-typed rows (the .vue parser bars casts). */
export function asFindResults(value: unknown): FindResults | null {
  return value !== null && typeof value === "object" && Array.isArray(Reflect.get(value, "rows")) ? (value as FindResults) : null;
}
