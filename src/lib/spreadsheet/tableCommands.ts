// The table commands (SPEC-133 "Creating a table", "Auto-expansion and
// calculated columns", "Header and totals rows", "Convert to range"). Every
// command mutates the passed model copy-on-write — `model.sheets`, the sheet,
// its `rows`/`tables`/`merges` arrays and the touched table are replaced,
// never mutated — marks the sheet `dirty` and the touched table `dirty`, and
// returns a result; the view records/commits undo around each ("Create
// table", "Expand table", "Fill table column", "Table options", "Table
// style", "Rename table", "Resize table", "Convert to range").
//
// Rules this module fixes where the spec leaves a choice:
// - Header names: a blank header (empty or spaces only) becomes the smallest
//   unused `Column{n}` (n ≥ 1, Excel's rule — a table created without headers
//   gets Column1..ColumnN); a duplicate (case-insensitive) gets the smallest
//   numeric suffix ≥ 2 unused by every header of the row (`Amount2`); a
//   number, date or boolean header becomes its displayed text; a header
//   formula becomes its value's text (Excel's "static text").
// - "Empty" (a row above/below, a cell auto-expansion or a fill may write) is
//   no value — or "" — and no formula; style-only cells are empty.
// - A protected sheet refuses every table command (PROTECTED_CELL_MESSAGE);
//   expansion and calculated fills simply do not happen when a cell they
//   would write is locked. A `locked` table (query/XML) never expands or
//   fills and refuses everything but setTableStyle and the style options of
//   setTableOptions — except a header-cell edit, whose column rename always
//   follows (the header cell must hold its column's name).
// - Totals: turning the row on the first time (no column has a totals
//   function or label) writes "Total" in the first column and SUBTOTAL sum
//   (every non-empty data cell a number, at least one) or count in the last
//   (a one-column table gets just the function); later times write each
//   column's kept function/label. Turning it off first reads the row's cells
//   back into the columns (a SUBTOTAL this module writes → its function, any
//   other formula → "custom", text → label, empty → none), so what the user
//   saw is what comes back.
// - Resize keeps the columns by sheet position; a table with a totals row
//   keeps it as the new range's last row, moving it there when the bottom
//   changes (refused when that row holds content).
import { getCell, type GridRange } from "./address";
import { firstLockedCell, PROTECTED_CELL_MESSAGE } from "./cells";
import type { DocumentTheme } from "../core/theme";
import { FORMAT_TOO_LARGE_MESSAGE, MAX_FORMAT_MATERIALISED } from "./formatting";
import { relativeFormulaKey, shiftFormula } from "./formula/references";
import { rangesIntersect } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type Sheet, type SheetTable, type SheetTableColumn, type SpreadsheetModel, type TotalsFunction } from "./model";
import { NAME_PROBLEM_MESSAGES, rewriteWorkbookFormulas, uniqueTableName, validateTableName, type NameProblem } from "./names";
import { formatCodeOf } from "./numberFormat/builtin";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES, type SpreadsheetLocale } from "./numberFormat/locale";
import { formatValue, plainText, workbookFormatContext } from "./numberFormat/render";
import { firstRowAtOrAfter } from "./restyle";
import { SheetEdit } from "./sheetEdit";
import { applyStructuralEdit, INSERT_PUSHES_DATA_MESSAGE, NOT_A_GRID_MESSAGE } from "./structure";
import { cellStyle, effectiveStyleId, internStyle, styleKey } from "./styles";
import { styleOverLook, tableCellLook } from "./tableStyles";
import {
  dataBody,
  expandStructuredReferences,
  formatStructured,
  qualifyStructuredReferences,
  removeTableReferences,
  renameColumnInFormula,
  renameTableInFormula,
  tableAt,
} from "./tables";

export const TABLES_OVERLAP_MESSAGE = "Tables can't overlap";
export const TABLE_ARRAY_MESSAGE = "Tables can't contain multi-cell array formulas";
export const RESIZE_TABLE_MESSAGE = "The header row must stay in the same row and the new range must overlap the table";
export const LOCKED_TABLE_MESSAGE = "This table is linked to external data — only its style can change";
export const TABLE_NOT_FOUND_MESSAGE = "That table no longer exists";
export const TOTALS_ROW_BLOCKED_MESSAGE = "The totals row can't move onto cells that hold data";

export const DEFAULT_TABLE_STYLE = "TableStyleMedium2";

export type TableResult = { ok: true } | { ok: false; message: string; problem?: NameProblem };
export type CreateTableResult = { ok: true; table: SheetTable } | { ok: false; message: string };

export type TotalsMenuId = Exclude<TotalsFunction, "custom"> | "none";
/** Excel's SUBTOTAL codes (101–110: hidden rows ignored) per totals function. */
export const SUBTOTAL_CODES: Readonly<Record<Exclude<TotalsFunction, "custom">, number>> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
};
/** The totals-cell menu, in order (`sheet-table-totals-<id>`). */
export const TOTALS_FUNCTIONS: readonly TotalsMenuId[] = ["none", "average", "count", "countNums", "max", "min", "sum", "stdDev", "var"];
export const TOTALS_FUNCTION_LABELS: Readonly<Record<TotalsMenuId, string>> = {
  none: "None",
  average: "Average",
  count: "Count",
  countNums: "Count Numbers",
  max: "Max",
  min: "Min",
  sum: "Sum",
  stdDev: "StdDev",
  var: "Var",
};

export interface TableOptionsPatch {
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  filterButton?: boolean;
}

// --- helpers -----------------------------------------------------------------------------------

const upper = (text: string): string => text.toUpperCase();

function isBlank(cell: Cell | undefined): boolean {
  return cell === undefined || (cell.formula === undefined && (cell.value === null || cell.value === ""));
}

/** Whether every cell of `row` in [left, right] is empty. */
function rowStripBlank(sheet: Sheet, row: number, left: number, right: number): boolean {
  const at = firstRowAtOrAfter(sheet.rows, row);
  const stored = sheet.rows[at];
  return stored?.index !== row || stored.cells.every((cell) => cell.col < left || cell.col > right || isBlank(cell));
}

function overlappingTable(sheet: Sheet, range: GridRange, exceptId?: string): SheetTable | undefined {
  return sheet.tables?.find((table) => table.id !== exceptId && rangesIntersect(table.ref, range));
}

/** Whether a multi-cell array formula reaches into `range`. */
function holdsMultiCellArray(sheet: Sheet, range: GridRange): boolean {
  return sheet.rows.some((row) =>
    row.cells.some((cell) => {
      const a = cell.arrayRange;
      return a !== undefined && (a.top !== a.bottom || a.left !== a.right) && rangesIntersect(a, range);
    }),
  );
}

/** A header cell's text: a string as is, any other value as displayed. */
function cellText(model: SpreadsheetModel, sheet: Sheet, row: number, col: number, locale: SpreadsheetLocale): string {
  const value = getCell(sheet, row, col)?.value ?? null;
  if (value === null || typeof value === "string") {
    return value ?? "";
  }
  const style = cellStyle(model, effectiveStyleId(sheet, row, col));
  return plainText(formatValue(value, formatCodeOf(style, locale), workbookFormatContext(model, locale))).trim();
}

function freeColumnName(taken: ReadonlySet<string>): string {
  for (let n = 1; ; n += 1) {
    if (!taken.has(`COLUMN${n}`)) {
      return `Column${n}`;
    }
  }
}

/** A typed header made a unique column name (see the module header). */
function uniqueColumnName(text: string, taken: ReadonlySet<string>): string {
  if (text.trim() === "") {
    return freeColumnName(taken);
  }
  if (!taken.has(upper(text))) {
    return text;
  }
  for (let n = 2; ; n += 1) {
    if (!taken.has(upper(`${text}${n}`))) {
      return `${text}${n}`;
    }
  }
}

/** A header row's texts as unique column names; `reserved` names are
 * avoided too. */
function normaliseHeaders(texts: readonly string[], reserved: readonly string[] = []): string[] {
  const literal = new Set([...texts.filter((text) => text.trim() !== "").map(upper), ...reserved.map(upper)]);
  const used = new Set(reserved.map(upper));
  return texts.map((text) => {
    const name = text.trim() !== "" && !used.has(upper(text)) ? text : uniqueColumnName(text, new Set([...literal, ...used]));
    used.add(upper(name));
    return name;
  });
}

/** Writes a header name into its cell unless it already holds exactly it. */
function writeHeader(edit: SheetEdit, row: number, col: number, name: string): void {
  const cell = edit.cell(row, col);
  if (cell?.value !== name || cell.formula !== undefined) {
    edit.change(row, col, { content: { value: name } });
  }
}

function clearStrip(edit: SheetEdit, range: GridRange): void {
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (edit.cell(row, col) !== undefined) {
        edit.change(row, col, { content: { value: null } });
      }
    }
  }
}

const nextColumnId = (columns: readonly SheetTableColumn[]): number => Math.max(columns.length, ...columns.map((column) => column.xlsxId ?? 0)) + 1;

interface Located {
  index: number;
  sheet: Sheet;
  table: SheetTable;
}

function locate(model: SpreadsheetModel, sheetId: string, tableId: string): Located | undefined {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  const table = sheet?.tables?.find((each) => each.id === tableId);
  return sheet && table ? { index, sheet, table } : undefined;
}

/** The refusal every edit of the table's definition shares, or null. */
function editRefusal(found: Located | undefined, styleOnly = false): string | null {
  if (!found) {
    return TABLE_NOT_FOUND_MESSAGE;
  }
  if (found.sheet.protection) {
    return PROTECTED_CELL_MESSAGE;
  }
  return found.table.locked && !styleOnly ? LOCKED_TABLE_MESSAGE : null;
}

/** Puts the sheet at `index` back with its new rows/merges and the table
 * replaced (null: removed), both dirty. */
function commit(model: SpreadsheetModel, index: number, tableId: string, table: SheetTable | null, patch: Partial<Pick<Sheet, "rows" | "merges">> = {}): void {
  const sheet = model.sheets[index]!;
  const tables = (sheet.tables ?? []).flatMap((each) => (each.id !== tableId ? [each] : table ? [{ ...table, dirty: true as const }] : []));
  const next: Sheet = { ...sheet, ...patch, tables, dirty: true };
  if (tables.length === 0) {
    delete next.tables;
  }
  model.sheets = model.sheets.map((each, i) => (i === index ? next : each));
}

function failed(message: string): { ok: false; message: string } {
  return { ok: false, message };
}

/** The row `table` (or a range) gains a header in: the row above when it is
 * empty across the columns and outside every table, else a row inserted at
 * `top` (SPEC-124). Returns the header row's index, or a refusal. */
function makeRowAbove(model: SpreadsheetModel, sheetId: string, range: GridRange): { row: number; inserted: boolean } | { message: string } {
  const sheet = model.sheets.find((each) => each.id === sheetId)!;
  const above: GridRange = { ...range, top: range.top - 1, bottom: range.top - 1 };
  if (range.top > 0 && rowStripBlank(sheet, above.top, range.left, range.right) && !overlappingTable(sheet, above)) {
    return { row: above.top, inserted: false };
  }
  const result = applyStructuralEdit(model, { kind: "insertRows", sheetId, at: range.top, count: 1 });
  return result.ok ? { row: range.top, inserted: true } : { message: result.message };
}

// --- creating a table -----------------------------------------------------------------------

/** Whether the range's first row looks like headers (SPEC-133 "Creating a
 * table"); `model` gives the styles for the bold rule (without it only the
 * value rule applies). */
export function detectHeaderRow(sheet: Sheet, range: GridRange, model?: Pick<SpreadsheetModel, "styles">): boolean {
  if (range.bottom <= range.top) {
    return false;
  }
  const bold = (row: number, col: number): boolean => model !== undefined && cellStyle(model, effectiveStyleId(sheet, row, col)).font?.bold === true;
  let secondNotText = false;
  let boldRule = true;
  for (let col = range.left; col <= range.right; col += 1) {
    const head = getCell(sheet, range.top, col)?.value;
    if (typeof head !== "string" || head === "") {
      return false;
    }
    secondNotText ||= typeof getCell(sheet, range.top + 1, col)?.value !== "string";
    boldRule &&= bold(range.top, col) && !bold(range.top + 1, col);
  }
  return secondNotText || (model !== undefined && boldRule);
}

const normalised = (range: GridRange): GridRange => ({
  top: Math.min(range.top, range.bottom),
  left: Math.min(range.left, range.right),
  bottom: Math.max(range.top, range.bottom),
  right: Math.max(range.left, range.right),
});

/** Insert ▸ Table (see the module header): "Create table". */
export function createTable(model: SpreadsheetModel, sheetId: string, range: GridRange, hasHeaders: boolean, locale: SpreadsheetLocale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]): CreateTableResult {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet || sheet.kind !== "grid") {
    return failed(NOT_A_GRID_MESSAGE);
  }
  if (sheet.protection) {
    return failed(PROTECTED_CELL_MESSAGE);
  }
  const given = normalised(range);
  // The final extent before any insert: a header row above, or a data row below.
  let extent = given;
  if (hasHeaders && given.top === given.bottom) {
    if (given.bottom + 1 >= MAX_ROWS) {
      return failed(INSERT_PUSHES_DATA_MESSAGE);
    }
    extent = { ...given, bottom: given.bottom + 1 };
  }
  if (overlappingTable(sheet, extent)) {
    return failed(TABLES_OVERLAP_MESSAGE);
  }
  if (holdsMultiCellArray(sheet, extent)) {
    return failed(TABLE_ARRAY_MESSAGE);
  }
  let ref = extent;
  if (!hasHeaders) {
    if (given.bottom + 1 >= MAX_ROWS) {
      return failed(INSERT_PUSHES_DATA_MESSAGE);
    }
    const header = makeRowAbove(model, sheetId, given);
    if ("message" in header) {
      return failed(header.message);
    }
    ref = { ...given, top: header.row, bottom: given.bottom + (header.inserted ? 1 : 0) };
  }
  const index = model.sheets.findIndex((each) => each.id === sheetId);
  const current = model.sheets[index]!;
  const width = ref.right - ref.left + 1;
  const texts = hasHeaders ? Array.from({ length: width }, (_, k) => cellText(model, current, ref.top, ref.left + k, locale)) : Array.from({ length: width }, () => "");
  const names = normaliseHeaders(texts);
  const edit = new SheetEdit(current);
  names.forEach((name, k) => writeHeader(edit, ref.top, ref.left + k, name));
  const table: SheetTable = {
    id: crypto.randomUUID(),
    name: uniqueTableName(model),
    ref,
    headerRow: true,
    totalsRow: false,
    columns: names.map((name, k) => ({ name, xlsxId: k + 1 })),
    style: { name: DEFAULT_TABLE_STYLE, showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
    autoFilter: {},
    dirty: true,
  };
  const merges = current.merges.filter((merge) => !rangesIntersect(merge, ref));
  model.sheets = model.sheets.map((each, i) =>
    i === index ? { ...current, rows: edit.rows(), merges: merges.length === current.merges.length ? current.merges : merges, tables: [...(current.tables ?? []), table], dirty: true } : each,
  );
  return { ok: true, table };
}

// --- auto-expansion and calculated columns -------------------------------------------------------

/** The table a typed entry at (row, col) would grow: directly below the last
 * data row of a table without a totals row, or directly right of the last
 * column within header..last data row. Locked tables never grow. */
function expansionTarget(sheet: Sheet, row: number, col: number): { table: SheetTable; down: boolean } | undefined {
  const tables = (sheet.tables ?? []).filter((table) => !table.locked);
  const below = tables.find((table) => !table.totalsRow && row === table.ref.bottom + 1 && col >= table.ref.left && col <= table.ref.right);
  if (below) {
    return { table: below, down: true };
  }
  const right = tables.find((table) => col === table.ref.right + 1 && row >= table.ref.top && row <= dataBody(table).bottom);
  return right ? { table: right, down: false } : undefined;
}

/** A typed formula about to be stored at (row, col): structured references
 * without a table name are qualified with the table there — or the one the
 * entry would grow — and `[#This Row]` becomes `@`; else unchanged. */
export function qualifyEntryFormula(model: SpreadsheetModel, sheetId: string, row: number, col: number, canonical: string): string {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  const table = sheet && (tableAt(sheet, row, col) ?? expansionTarget(sheet, row, col)?.table);
  return table ? qualifyStructuredReferences(canonical, table.name) : canonical;
}

/** After a typed entry's commit (never paste, fill, undo or a formula
 * result): grows the table the entry is directly below or right of by one
 * row or column — "Expand table". Whether it expanded. */
export function expandTableForEntry(model: SpreadsheetModel, sheetId: string, row: number, col: number, locale: SpreadsheetLocale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]): boolean {
  const index = model.sheets.findIndex((each) => each.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || sheet.kind !== "grid" || row >= MAX_ROWS || col >= MAX_COLS || isBlank(getCell(sheet, row, col))) {
    return false;
  }
  const target = expansionTarget(sheet, row, col);
  if (!target) {
    return false;
  }
  const { table, down } = target;
  const ref = down ? { ...table.ref, bottom: row } : { ...table.ref, right: col };
  const strip = down ? { ...ref, top: row } : { ...ref, left: col };
  if (overlappingTable(sheet, strip, table.id) || firstLockedCell(model, index, [strip])) {
    return false;
  }
  const edit = new SheetEdit(sheet);
  let columns = table.columns;
  if (down) {
    const first = dataBody(table).top;
    table.columns.forEach((column, k) => {
      if (column.calculatedFormula !== undefined && isBlank(edit.cell(row, ref.left + k))) {
        edit.change(row, ref.left + k, { content: { value: null, formula: shiftFormula(column.calculatedFormula, row - first, 0) } });
      }
    });
  } else {
    const taken = new Set(table.columns.map((column) => upper(column.name)));
    const name = table.headerRow ? uniqueColumnName(cellText(model, sheet, ref.top, col, locale), taken) : freeColumnName(taken);
    if (table.headerRow) {
      writeHeader(edit, ref.top, col, name);
    }
    columns = [...columns, { name, xlsxId: nextColumnId(columns) }];
  }
  commit(model, index, table.id, { ...table, ref, columns }, { rows: edit.rows() });
  return true;
}

/** After a formula's commit into a table's data cell: when the column's
 * other data cells are all empty or all hold (shifted) its current
 * calculatedFormula, fills it into every data row and makes it the column's
 * calculatedFormula — "Fill table column". Whether anything changed. */
export function fillCalculatedColumn(model: SpreadsheetModel, sheetId: string, row: number, col: number): boolean {
  const index = model.sheets.findIndex((each) => each.id === sheetId);
  const sheet = model.sheets[index];
  const cell = sheet && getCell(sheet, row, col);
  const table = sheet && tableAt(sheet, row, col);
  if (!sheet || !table || table.locked || cell?.formula === undefined || cell.formulaUnparsed || cell.arrayRange) {
    return false;
  }
  const body = dataBody(table);
  const k = col - table.ref.left;
  const column = table.columns[k];
  if (!column || row < body.top || row > body.bottom) {
    return false;
  }
  const strip = { ...body, left: col, right: col };
  if (firstLockedCell(model, index, [strip]) || holdsMultiCellArray(sheet, strip)) {
    return false;
  }
  const calculatedKey = column.calculatedFormula === undefined ? undefined : relativeFormulaKey(column.calculatedFormula, { row: body.top, col });
  let allEmpty = true;
  let allCalculated = calculatedKey !== undefined;
  for (let r = body.top; r <= body.bottom && (allEmpty || allCalculated); r += 1) {
    const other = r === row ? undefined : getCell(sheet, r, col);
    if (r === row || isBlank(other)) {
      allCalculated &&= r === row;
      continue;
    }
    allEmpty = false;
    allCalculated &&= other!.formula !== undefined && !other!.formulaUnparsed && relativeFormulaKey(other!.formula, { row: r, col }) === calculatedKey;
  }
  if (!allEmpty && !allCalculated) {
    return false;
  }
  const edit = new SheetEdit(sheet);
  for (let r = body.top; r <= body.bottom; r += 1) {
    const formula = shiftFormula(cell.formula, r - row, 0);
    const existing = getCell(sheet, r, col);
    if (r !== row && (existing?.formula !== formula || existing.formulaUnparsed)) {
      edit.change(r, col, { content: { value: null, formula } });
    }
  }
  const calculatedFormula = shiftFormula(cell.formula, body.top - row, 0);
  const rows = edit.rows();
  if (rows === sheet.rows && calculatedFormula === column.calculatedFormula) {
    return false;
  }
  const columns = table.columns.map((each, i) => (i === k ? { ...each, calculatedFormula } : each));
  commit(model, index, table.id, { ...table, columns }, { rows });
  return true;
}

// --- totals and header rows ------------------------------------------------------------------

/** The formula a totals function writes (`custom`: the column's own). */
function totalsFormulaOf(table: SheetTable, column: SheetTableColumn, fn: TotalsFunction): string | undefined {
  if (fn === "custom") {
    return column.totalsFormula;
  }
  return `SUBTOTAL(${SUBTOTAL_CODES[fn]},${formatStructured({ table: table.name, items: [], thisRow: false, columns: { first: column.name } })})`;
}

/** What a column's totals cell holds, or null for nothing. */
function totalsContent(table: SheetTable, column: SheetTableColumn): { value: CellValue; formula?: string } | null {
  if (column.totalsFunction !== undefined) {
    const formula = totalsFormulaOf(table, column, column.totalsFunction);
    return formula === undefined ? null : { value: null, formula };
  }
  return column.totalsLabel === undefined ? null : { value: column.totalsLabel };
}

/** The columns with their totals fields read back from the totals row's
 * cells (see the module header). */
function columnsFromTotalsRow(sheet: Sheet, table: SheetTable): SheetTableColumn[] {
  return table.columns.map((column, k) => {
    const cell = getCell(sheet, table.ref.bottom, table.ref.left + k);
    const { totalsFunction: _fn, totalsFormula: _formula, totalsLabel: _label, ...rest } = column;
    if (isBlank(cell)) {
      return rest;
    }
    if (cell!.formula !== undefined && !cell!.formulaUnparsed) {
      const fn = (Object.keys(SUBTOTAL_CODES) as Exclude<TotalsFunction, "custom">[]).find((each) => totalsFormulaOf(table, column, each) === cell!.formula);
      return fn ? { ...rest, totalsFunction: fn } : { ...rest, totalsFunction: "custom", totalsFormula: cell!.formula };
    }
    return typeof cell!.value === "string" ? { ...rest, totalsLabel: cell!.value } : column;
  });
}

/** Whether every non-empty data cell of the column is a number (and one is). */
function numericColumn(sheet: Sheet, body: GridRange, col: number): boolean {
  let numbers = 0;
  for (let r = firstRowAtOrAfter(sheet.rows, body.top); r < sheet.rows.length && sheet.rows[r]!.index <= body.bottom; r += 1) {
    const cell = getCell(sheet, sheet.rows[r]!.index, col);
    if (isBlank(cell)) {
      continue;
    }
    if (typeof cell!.value !== "number") {
      return false;
    }
    numbers += 1;
  }
  return numbers > 0;
}

function writeTotals(edit: SheetEdit, table: SheetTable, row: number): void {
  table.columns.forEach((column, k) => {
    const content = totalsContent(table, column);
    if (content) {
      edit.change(row, table.ref.left + k, { content });
    }
  });
}

/** Table options ▸ Total row (see the module header). */
export function setTotalsRow(model: SpreadsheetModel, sheetId: string, tableId: string, on: boolean): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { sheet, table } = found!;
  if (table.totalsRow === on) {
    return { ok: true };
  }
  if (!on) {
    const edit = new SheetEdit(sheet);
    clearStrip(edit, { ...table.ref, top: table.ref.bottom });
    const next = { ...table, ref: { ...table.ref, bottom: table.ref.bottom - 1 }, totalsRow: false, columns: columnsFromTotalsRow(sheet, table) };
    commit(model, found!.index, tableId, next, { rows: edit.rows() });
    return { ok: true };
  }
  const row = table.ref.bottom + 1;
  if (row >= MAX_ROWS) {
    return failed(INSERT_PUSHES_DATA_MESSAGE);
  }
  if (!rowStripBlank(sheet, row, table.ref.left, table.ref.right) || overlappingTable(sheet, { ...table.ref, top: row, bottom: row })) {
    const result = applyStructuralEdit(model, { kind: "insertRows", sheetId, at: row, count: 1 });
    if (!result.ok) {
      return failed(result.message);
    }
  }
  const after = locate(model, sheetId, tableId)!;
  let columns = after.table.columns;
  if (!columns.some((column) => column.totalsFunction !== undefined || column.totalsLabel !== undefined)) {
    const last = columns.length - 1;
    const fn = numericColumn(after.sheet, dataBody(table), table.ref.left + last) ? "sum" : "count";
    columns = columns.map((column, k) => (k === last ? { ...column, totalsFunction: fn } : k === 0 ? { ...column, totalsLabel: "Total" } : column));
  }
  const next: SheetTable = { ...after.table, ref: { ...table.ref, bottom: row }, totalsRow: true, columns };
  const edit = new SheetEdit(after.sheet);
  writeTotals(edit, next, row);
  commit(model, after.index, tableId, next, { rows: edit.rows() });
  return { ok: true };
}

/** The totals-cell menu: writes `SUBTOTAL(code,T[Col])`, the column's custom
 * formula, or (none) clears the cell. */
export function setTotalsFunction(model: SpreadsheetModel, sheetId: string, tableId: string, columnIndex: number, fn: TotalsFunction | "none"): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { sheet, table, index } = found!;
  const column = table.columns[columnIndex];
  if (!column) {
    return failed(TABLE_NOT_FOUND_MESSAGE);
  }
  const { totalsFunction: _fn, totalsLabel: _label, ...rest } = column;
  const nextColumn: SheetTableColumn = fn === "none" ? rest : { ...rest, totalsFunction: fn };
  if (fn !== "custom" && fn !== "none") {
    delete nextColumn.totalsFormula;
  }
  const next = { ...table, columns: table.columns.map((each, i) => (i === columnIndex ? nextColumn : each)) };
  const edit = new SheetEdit(sheet);
  if (table.totalsRow) {
    const content = totalsContent(next, nextColumn);
    edit.change(table.ref.bottom, table.ref.left + columnIndex, { content: content ?? { value: null } });
  }
  commit(model, index, tableId, next, { rows: edit.rows() });
  return { ok: true };
}

/** Table options ▸ Header row: off clears the header cells and moves
 * `ref.top` down (the names stay); on writes them back into the row above,
 * found or made as in createTable. */
export function setHeaderRow(model: SpreadsheetModel, sheetId: string, tableId: string, on: boolean): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { table } = found!;
  if (table.headerRow === on) {
    return { ok: true };
  }
  if (!on) {
    const edit = new SheetEdit(found!.sheet);
    clearStrip(edit, { ...table.ref, bottom: table.ref.top });
    commit(model, found!.index, tableId, { ...table, ref: { ...table.ref, top: table.ref.top + 1 }, headerRow: false }, { rows: edit.rows() });
    return { ok: true };
  }
  const header = makeRowAbove(model, sheetId, table.ref);
  if ("message" in header) {
    return failed(header.message);
  }
  // Computed from the old ref: the same whether or not the insert moved the table.
  const ref = { ...table.ref, top: header.row, bottom: table.ref.bottom + (header.inserted ? 1 : 0) };
  const after = locate(model, sheetId, tableId)!;
  const edit = new SheetEdit(after.sheet);
  after.table.columns.forEach((column, k) => writeHeader(edit, ref.top, ref.left + k, column.name));
  commit(model, after.index, tableId, { ...after.table, ref, headerRow: true }, { rows: edit.rows() });
  return { ok: true };
}

// --- options, style, names --------------------------------------------------------------------

/** The design bar's check boxes — "Table options". A locked table takes the
 * style options only. */
export function setTableOptions(model: SpreadsheetModel, sheetId: string, tableId: string, patch: TableOptionsPatch): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found, patch.filterButton === undefined);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { table } = found!;
  const { filterButton, ...options } = patch;
  const style = { ...table.style, ...Object.fromEntries(Object.entries(options).filter(([, value]) => value !== undefined)) };
  const next: SheetTable = { ...table, style };
  if (filterButton === true && !table.autoFilter) {
    next.autoFilter = {};
  } else if (filterButton === false) {
    delete next.autoFilter;
  }
  const sameStyle = (Object.keys(style) as (keyof typeof style)[]).every((key) => style[key] === table.style[key]);
  if (sameStyle && (next.autoFilter === undefined) === (table.autoFilter === undefined)) {
    return { ok: true };
  }
  commit(model, found!.index, tableId, next);
  return { ok: true };
}

/** The style gallery — "Table style"; `undefined` is None. */
export function setTableStyle(model: SpreadsheetModel, sheetId: string, tableId: string, name: string | undefined): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found, true);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { table } = found!;
  if (table.style.name === name) {
    return { ok: true };
  }
  const style = { ...table.style };
  if (name === undefined) {
    delete style.name;
  } else {
    style.name = name;
  }
  commit(model, found!.index, tableId, { ...table, style });
  return { ok: true };
}

/** The design bar's name field — "Rename table": validated, then every
 * formula and name using the old name is rewritten. */
export function renameTable(model: SpreadsheetModel, tableId: string, name: string): TableResult {
  const sheet = model.sheets.find((each) => each.tables?.some((table) => table.id === tableId));
  const found = sheet && locate(model, sheet.id, tableId);
  const refusal = editRefusal(found);
  if (refusal !== null) {
    return failed(refusal);
  }
  const old = found!.table.name;
  if (old === name) {
    return { ok: true };
  }
  const problem = validateTableName(model, name, tableId);
  if (problem !== null) {
    return { ok: false, message: NAME_PROBLEM_MESSAGES[problem], problem };
  }
  rewriteWorkbookFormulas(model, (text) => renameTableInFormula(text, old, name));
  const after = locate(model, sheet!.id, tableId)!;
  commit(model, after.index, tableId, { ...after.table, name });
  return { ok: true };
}

/** The header cell a commit at (row, col) edited, when it is one: its table,
 * column and the cell's text (what renameTableColumn takes). */
export function headerEditColumn(
  model: SpreadsheetModel,
  sheetId: string,
  row: number,
  col: number,
  locale: SpreadsheetLocale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE],
): { tableId: string; columnIndex: number; text: string } | undefined {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  const table = sheet && tableAt(sheet, row, col);
  if (!sheet || !table?.headerRow || row !== table.ref.top) {
    return undefined;
  }
  return { tableId: table.id, columnIndex: col - table.ref.left, text: cellText(model, sheet, row, col, locale) };
}

/** A header cell's edit renames its column (blank → Column{n}, a duplicate
 * suffixed) and every reference to it; the header cell is left holding the
 * name as a string. */
export function renameTableColumn(model: SpreadsheetModel, sheetId: string, tableId: string, columnIndex: number, name: string): TableResult {
  const found = locate(model, sheetId, tableId);
  const column = found?.table.columns[columnIndex];
  if (!found || !column) {
    return failed(TABLE_NOT_FOUND_MESSAGE);
  }
  const { table } = found;
  const col = table.ref.left + columnIndex;
  if (table.headerRow && firstLockedCell(model, found.index, [{ top: table.ref.top, bottom: table.ref.top, left: col, right: col }])) {
    return failed(PROTECTED_CELL_MESSAGE);
  }
  const taken = new Set(table.columns.filter((_, i) => i !== columnIndex).map((each) => upper(each.name)));
  const wanted = uniqueColumnName(name, taken);
  if (wanted !== column.name) {
    rewriteWorkbookFormulas(model, (text) => renameColumnInFormula(text, table.name, column.name, wanted));
  }
  const after = locate(model, sheetId, tableId)!;
  const edit = new SheetEdit(after.sheet);
  if (table.headerRow) {
    writeHeader(edit, table.ref.top, col, wanted);
  }
  const rows = edit.rows();
  if (wanted === column.name && rows === after.sheet.rows) {
    return { ok: true };
  }
  const columns = after.table.columns.map((each, i) => (i === columnIndex ? { ...each, name: wanted } : each));
  commit(model, after.index, tableId, { ...after.table, columns }, { rows });
  return { ok: true };
}

// --- resize and convert ----------------------------------------------------------------------

/** The Resize Table dialog — "Resize table" (see the module header). */
export function resizeTable(
  model: SpreadsheetModel,
  sheetId: string,
  tableId: string,
  range: GridRange,
  locale: SpreadsheetLocale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE],
): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { sheet, table } = found!;
  const ref = normalised(range);
  const minRows = 1 + (table.headerRow ? 1 : 0) + (table.totalsRow ? 1 : 0);
  if (ref.top !== table.ref.top || !rangesIntersect(ref, table.ref) || ref.bottom - ref.top + 1 < minRows) {
    return failed(RESIZE_TABLE_MESSAGE);
  }
  if (overlappingTable(sheet, ref, tableId)) {
    return failed(TABLES_OVERLAP_MESSAGE);
  }
  if (holdsMultiCellArray(sheet, ref)) {
    return failed(TABLE_ARRAY_MESSAGE);
  }
  const movesTotals = table.totalsRow && ref.bottom !== table.ref.bottom;
  if (movesTotals && !rowStripBlank(sheet, ref.bottom, ref.left, ref.right)) {
    return failed(TOTALS_ROW_BLOCKED_MESSAGE);
  }
  // References to the columns leaving become #REF!.
  const removed = table.columns.filter((_, k) => table.ref.left + k < ref.left || table.ref.left + k > ref.right);
  if (removed.length > 0) {
    rewriteWorkbookFormulas(model, (text) => removed.reduce((out, column) => removeTableReferences(out, table.name, column.name), text));
  }
  const after = locate(model, sheetId, tableId)!;
  const edit = new SheetEdit(after.sheet);
  let columns = after.table.columns;
  if (movesTotals) {
    columns = columnsFromTotalsRow(after.sheet, after.table);
    clearStrip(edit, { ...table.ref, top: table.ref.bottom });
  }
  const kept = columns.filter((_, k) => table.ref.left + k >= ref.left && table.ref.left + k <= ref.right);
  const added: SheetTableColumn[] = [];
  const keptAt = (col: number): SheetTableColumn | undefined => (col >= table.ref.left && col <= table.ref.right ? columns[col - table.ref.left] : undefined);
  const newCols = Array.from({ length: ref.right - ref.left + 1 }, (_, k) => ref.left + k).filter((col) => !keptAt(col));
  const texts = newCols.map((col) => (table.headerRow ? cellText(model, after.sheet, ref.top, col, locale) : ""));
  const names = normaliseHeaders(texts, kept.map((column) => column.name));
  let nextId = nextColumnId(columns);
  const ordered = Array.from({ length: ref.right - ref.left + 1 }, (_, k): SheetTableColumn => {
    const col = ref.left + k;
    const existing = keptAt(col);
    if (existing) {
      return existing;
    }
    const name = names[newCols.indexOf(col)]!;
    if (table.headerRow) {
      writeHeader(edit, ref.top, col, name);
    }
    const column = { name, xlsxId: nextId };
    nextId += 1;
    added.push(column);
    return column;
  });
  const next: SheetTable = { ...after.table, ref, columns: ordered };
  if (movesTotals) {
    writeTotals(edit, next, ref.bottom);
  }
  const merges = after.sheet.merges.filter((merge) => !rangesIntersect(merge, ref));
  commit(model, after.index, tableId, next, { rows: edit.rows(), ...(merges.length === after.sheet.merges.length ? {} : { merges }) });
  return { ok: true };
}

/** Every formula of the workbook rewritten with its host cell: cells, names
 * (host: A1 of their scope sheet, else the first sheet) and every other
 * table's calculated (first data row) and totals (totals row) formulas. */
function rewriteHostedFormulas(model: SpreadsheetModel, skipTableId: string, rewrite: (text: string, host: { sheetId: string; row: number; col: number }) => string): void {
  const firstSheetId = model.sheets[0]?.id ?? "";
  model.sheets = model.sheets.map((sheet) => {
    let changed = false;
    const rows = sheet.rows.map((row) => {
      let cellsChanged = false;
      const cells = row.cells.map((cell) => {
        if (cell.formula === undefined || cell.formulaUnparsed) {
          return cell;
        }
        const formula = rewrite(cell.formula, { sheetId: sheet.id, row: row.index, col: cell.col });
        cellsChanged ||= formula !== cell.formula;
        return formula === cell.formula ? cell : { ...cell, formula };
      });
      changed ||= cellsChanged;
      return cellsChanged ? { ...row, cells } : row;
    });
    const tables = sheet.tables?.map((table) => {
      if (table.id === skipTableId) {
        return table;
      }
      let tableChanged = false;
      const columns = table.columns.map((column, k) => {
        const col = table.ref.left + k;
        const next = { ...column };
        if (column.calculatedFormula !== undefined) {
          next.calculatedFormula = rewrite(column.calculatedFormula, { sheetId: sheet.id, row: dataBody(table).top, col });
        }
        if (column.totalsFormula !== undefined) {
          next.totalsFormula = rewrite(column.totalsFormula, { sheetId: sheet.id, row: table.ref.bottom, col });
        }
        const same = next.calculatedFormula === column.calculatedFormula && next.totalsFormula === column.totalsFormula;
        tableChanged ||= !same;
        return same ? column : next;
      });
      changed ||= tableChanged;
      return tableChanged ? { ...table, columns, dirty: true as const } : table;
    });
    return changed ? { ...sheet, rows, ...(tables ? { tables } : {}), dirty: true as const } : sheet;
  });
  let namesChanged = false;
  const names = model.names.map((name) => {
    if (name.formulaUnparsed) {
      return name;
    }
    const formula = rewrite(name.formula, { sheetId: name.scopeSheetId ?? firstSheetId, row: 0, col: 0 });
    namesChanged ||= formula !== name.formula;
    return formula === name.formula ? name : { ...name, formula };
  });
  if (namesChanged) {
    model.names = names;
    model.namesDirty = true;
  }
}

/** Design ▸ Convert to Range — "Convert to range": each cell's table look
 * baked beneath its own style, every reference to the table rewritten to A1
 * (this-row `$C7`, the rest `$C$2:$C$10`), the table removed. */
export function convertTableToRange(model: SpreadsheetModel, sheetId: string, tableId: string, theme: DocumentTheme): TableResult {
  const found = locate(model, sheetId, tableId);
  const refusal = editRefusal(found);
  if (refusal !== null) {
    return failed(refusal);
  }
  const { sheet, table, index } = found!;
  const { ref } = table;
  if ((ref.bottom - ref.top + 1) * (ref.right - ref.left + 1) > MAX_FORMAT_MATERIALISED) {
    return failed(FORMAT_TOO_LARGE_MESSAGE);
  }
  const edit = new SheetEdit(sheet);
  const normal = model.styles[0];
  for (let row = ref.top; row <= ref.bottom; row += 1) {
    for (let col = ref.left; col <= ref.right; col += 1) {
      const look = tableCellLook(sheet, table, row, col, theme);
      if (!look) {
        continue;
      }
      const own = cellStyle(model, effectiveStyleId(sheet, row, col));
      const baked = styleOverLook(look, own, normal);
      if (styleKey(baked) !== styleKey(own)) {
        edit.change(row, col, { style: internStyle(model, baked) });
      }
    }
  }
  const rows = edit.rows();
  model.sheets = model.sheets.map((each, i) => (i === index ? { ...each, rows } : each));
  rewriteHostedFormulas(model, tableId, (text, host) => expandStructuredReferences(model, text, host, table.name));
  const at = model.sheets.findIndex((each) => each.id === sheetId);
  commit(model, at, tableId, null);
  return { ok: true };
}
