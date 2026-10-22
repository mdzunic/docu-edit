// The table design surfaces' fixed parts (SPEC-133 "View wiring"), framework-
// free for useSheetTables and its specs: the undo labels, the ODS notice, the
// design bar's checkboxes (the style ones a locked table keeps) and the
// dialogs' range text — shown absolute (`$A$1:$D$5`), read with a leading
// "=", `$` signs or the dialog's own sheet prefix.
import { colToLetters, parseRange, type GridRange } from "../../lib/spreadsheet/address";
import { parseSheetPrefix, sheetPrefixOf } from "../../lib/spreadsheet/formula/tokens";
import type { TotalsMenuId } from "../../lib/spreadsheet/tableCommands";

export const TABLE_UNDO_LABELS = {
  create: "Create table",
  rename: "Rename table",
  options: "Table options",
  style: "Table style",
  resize: "Resize table",
  convert: "Convert to range",
  expand: "Expand table",
  fill: "Fill table column",
} as const;

export const ODS_TABLE_NOTICE = "OpenDocument has no tables — this saves as a named range with the table's formatting applied to its cells";
export const TABLE_RANGE_MESSAGE = "Enter a cell range on this sheet, such as $A$1:$D$5";

export type TableOption = "headerRow" | "totalsRow" | "bandedRows" | "bandedColumns" | "firstColumn" | "lastColumn" | "filterButton";
export type TotalsChoice = TotalsMenuId;

/** The style options, which a locked table keeps, as setTableOptions' fields. */
export const STYLE_OPTIONS: Partial<Record<TableOption, "showRowStripes" | "showColumnStripes" | "showFirstColumn" | "showLastColumn">> = {
  bandedRows: "showRowStripes",
  bandedColumns: "showColumnStripes",
  firstColumn: "showFirstColumn",
  lastColumn: "showLastColumn",
};

export const OPTION_ROWS: readonly { option: TableOption; label: string; testid: string }[] = [
  { option: "headerRow", label: "Header row", testid: "sheet-table-option-header-row" },
  { option: "totalsRow", label: "Total row", testid: "sheet-table-option-totals-row" },
  { option: "bandedRows", label: "Banded rows", testid: "sheet-table-option-banded-rows" },
  { option: "bandedColumns", label: "Banded columns", testid: "sheet-table-option-banded-columns" },
  { option: "firstColumn", label: "First column", testid: "sheet-table-option-first-column" },
  { option: "lastColumn", label: "Last column", testid: "sheet-table-option-last-column" },
  { option: "filterButton", label: "Filter button", testid: "sheet-table-option-filter-button" },
];

const corner = (row: number, col: number): string => `$${colToLetters(col)}$${row + 1}`;

/** A range as the dialogs show it: absolute, `$A$1:$D$5`. */
export function tableRangeText(range: GridRange): string {
  const start = corner(range.top, range.left);
  return range.top === range.bottom && range.left === range.right ? start : `${start}:${corner(range.bottom, range.right)}`;
}

/** A dialog's range text on sheet `sheetName` (a leading "=", `$` signs and
 * that sheet's own prefix allowed); null when unreadable or on another sheet. */
export function parseTableRange(text: string, sheetName: string): GridRange | null {
  let body = text.trim().replace(/^=\s*/, "");
  const prefix = sheetPrefixOf(body);
  if (prefix !== "") {
    if (parseSheetPrefix(prefix).sheet.toUpperCase() !== sheetName.toUpperCase()) {
      return null;
    }
    body = body.slice(prefix.length);
  }
  return parseRange(body.replaceAll("$", ""));
}
