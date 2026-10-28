// The workbook generation API's shared core (SPEC-140) — what deckSlides.ts
// is to decks: the accumulated state both front ends drive, the fluent
// `WorkbookBuilder` (workbook.ts) and the JSON description
// (workbookDescription.ts), so the two produce identical models and report
// identical paths for identical mistakes.
//
// Every call validates all of its arguments before it records anything, and
// names the offending field by its JSON-description path, counting sheets
// and calls in order: `sheets[0].cells.B5.formula`,
// `sheets[0].data[1].rows[3][2]`, `sheets[1].styles[0].style.font.sizePt`,
// `styles.header.fill`. Fluent calls count as the description's entries
// do: setRow and setRows are `data[j]`, style and numberFormat `styles[k]`,
// columnWidth `columns[k]`, rowHeight `rows[k]`, merge `merges[k]`.
//
// Nothing here touches the model until `materialise`, which builds a fresh
// one each time in the pinned order (SPEC-140 "Materialisation"):
//  1. createBlankWorkbook() with the title, format, date system and the
//     builder's sheets (none added → its one empty "Sheet1");
//  2. per sheet, the accumulated cells sorted into rows, the last write to a
//     cell winning (dates were converted, and parseText applied, on entry —
//     the builder's locale and date system are fixed at construction);
//  3. the style calls in call order through SPEC-123's applyStylePatch — a
//     cell's own style/numberFormat (and a date's default format) counting
//     as a style call made when the cell was set. A run of consecutive
//     cell-own calls is applied as one applyStylePatch per patch over the
//     fewest rectangles, never one copy-on-write command per cell (patches
//     in order of first use; a cell taking them out of that order starts a
//     new batch, so every cell sees its patches in call order);
//  4. merges, through SPEC-124's mergeCells(…, "cells"), refused when a
//     value outside the top-left cell would go;
//  5. column widths and row heights (bands.ts's resize commands), frozen
//     panes (freezePanes) and the sheet options;
//  6. whole-workbook validation: every sheet a formula names exists, a
//     visible sheet remains, and the stored cells stay within
//     MAX_WORKBOOK_CELLS;
//  7. SPEC-120's recalculate over every formula cell — a cycle is refused at
//     its first cell, as is a cell the engine cannot compute;
//  8. a structural self-check against SPEC-116's invariants.
//
// Imports lib/spreadsheet, lib/core and generate's product-neutral modules
// only (workbookBoundary.spec.ts); lib/spreadsheet never imports this.
import { colToLetters, formatA1, formatRange, lettersToCol, parseA1, parseRange, type GridRange } from "../spreadsheet/address";
import { MAX_COLUMN_WIDTH_CHARS, MAX_ROW_HEIGHT_PT, resizeColumns, resizeRows, freezePanes } from "../spreadsheet/bands";
import { isEmptyRow } from "../spreadsheet/cells";
import { inferCsvField } from "../spreadsheet/csv/infer";
import { applyStylePatch, type StylePatch } from "../spreadsheet/formatting";
import type { FormulaNode } from "../spreadsheet/formula/ast";
import { cellKey, type CellKey } from "../spreadsheet/formula/dependencies";
import { FormulaSyntaxError, parseFormula } from "../spreadsheet/formula/parse";
import { printFormula, quoteSheetName } from "../spreadsheet/formula/print";
import { recalculate, uncomputableReasonOf } from "../spreadsheet/formula/recalc";
import { getFunction } from "../spreadsheet/formula/registry";
import { xlsxWidthToPx } from "../spreadsheet/geometry";
import { mergeCells } from "../spreadsheet/mergeCells";
import { rangesIntersect } from "../spreadsheet/merges";
import {
  ERROR_CODES,
  MAX_COLS,
  MAX_ROWS,
  MAX_WORKBOOK_CELLS,
  type Cell,
  type CellValue,
  type ErrorCode,
  type Row,
  type Sheet,
  type SpreadsheetModel,
} from "../spreadsheet/model";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../spreadsheet/numberFormat/locale";
import { dateToSerial } from "../spreadsheet/serial";
import { createBlankWorkbook, emptySheet, LAST_VISIBLE_SHEET_MESSAGE, validateSheetName } from "../spreadsheet/sheets";
import { fontMaxDigitWidth } from "../spreadsheet/styles";
import { WORKBOOK_TOO_LARGE_MESSAGE } from "../spreadsheet/errors";
import { DEFAULT_THEME } from "../core/theme";
import { DocumentGenerationError } from "./errors";
import {
  expectArray,
  expectBoolean,
  expectEnum,
  expectHexColor,
  expectIntInRange,
  expectNonNegativeNumber,
  expectPlainObject,
  expectString,
  rejectUnknownKeys,
} from "./validate";
import { expectNumberFormat, styleSpecPatches, validateStyleSpec, type CellStyleSpec } from "./workbookStyles";
// SPEC-120/127's function library registers itself on import — evaluate.ts
// loads it too; named here because the entry check reads the registry.
import "../spreadsheet/formula/functions";

export type WorkbookOutputFormat = "xlsx" | "ods";
export const WORKBOOK_OUTPUT_FORMATS: readonly WorkbookOutputFormat[] = ["xlsx", "ods"];

export interface WorkbookBuilderOptions {
  title?: string;
  /** The workbook's own format — what toBytes()/fileName() use when called
   * without one. Defaults to "xlsx". */
  format?: WorkbookOutputFormat;
  /** Defaults to "1900". */
  dateSystem?: "1900" | "1904";
  /** The locale `parseText` reads fields in and formulas evaluate in (their
   * text is always en-US syntax). Defaults to "en-US". */
  locale?: SpreadsheetLocaleTag;
}

/** One cell's input (SPEC-140 "Cell input"). */
export type CellInput = string | number | boolean | Date | null | CellInputObject;

export interface CellInputObject {
  value?: string | number | boolean | null;
  /** "YYYY-MM-DD[THH:MM[:SS[.sss]]]", no UTC offset. */
  date?: string;
  error?: ErrorCode;
  /** SPEC-119 canonical text; a leading "=" is accepted. */
  formula?: string;
  numberFormat?: string;
  /** A spec, or the name of a style defineStyle registered. */
  style?: CellStyleSpec | string;
}

export interface AddSheetOptions {
  /** `#rrggbb`. */
  tabColor?: string;
  hidden?: boolean;
  /** Defaults to true. */
  showGridLines?: boolean;
}

export interface SetRowsOptions {
  /** Route every string through SPEC-126's CSV inference (never a formula). */
  parseText?: boolean;
}

export type ColumnWidthSpec = { chars: number } | { px: number };
export type RowHeightSpec = { pt: number } | { px: number };

export interface FreezeSpec {
  rows?: number;
  cols?: number;
}

export const OPTION_KEYS = ["title", "format", "dateSystem", "locale"] as const;
export const SHEET_OPTION_KEYS = ["tabColor", "hidden", "showGridLines"] as const;
export const CELL_INPUT_KEYS = ["value", "date", "error", "formula", "numberFormat", "style"] as const;
const CONTENT_KEYS = ["value", "date", "error", "formula"] as const;

/** Excel's cell text limit. */
export const MAX_CELL_TEXT_LENGTH = 32_767;
/** The default date and date-time formats (SPEC-140 "Cell input"). */
export const DATE_FORMAT = "yyyy-mm-dd";
export const DATE_TIME_FORMAT = "yyyy-mm-dd hh:mm";
/** `{ px }` row heights convert at 0.75 pt per px. */
const PT_PER_PX = 0.75;
/** The maximum digit width column widths convert with: a generated
 * workbook's default font is always the blank workbook's Calibri 11 (7 px). */
const MDW = fontMaxDigitWidth(createBlankWorkbook().styles[0]!.font, DEFAULT_THEME);
const MAX_COLUMN_WIDTH_PX = xlsxWidthToPx(MAX_COLUMN_WIDTH_CHARS, MDW);
const MS_PER_DAY = 86_400_000;

export const NO_SHEET_MESSAGE = "no sheet — call addSheet first";

// --- the accumulated state ---------------------------------------------------------------

/** One accumulated cell: its content as of the last write. */
interface CellEntry {
  value: CellValue;
  formula?: string;
  /** The sheets the formula names (upper-cased), checked at output. */
  sheets?: string[];
  /** The write it came from, as the path of that call's value. */
  path: string;
}

/** One applyStylePatch call; `cell` marks a cell's own (coalescible) call. */
interface StyleCall {
  ranges: GridRange[];
  patch: StylePatch;
  path: string;
  cell?: true;
}

interface MergeCall {
  range: GridRange;
  path: string;
}

interface ColumnCall {
  start: number;
  end: number;
  px: number;
  chars?: number;
  path: string;
}

interface RowCall {
  start: number;
  end: number;
  pt: number;
  path: string;
}

interface PlanSheet {
  name: string;
  path: string;
  tabColor?: string;
  hidden: boolean;
  showGridLines: boolean;
  /** row × MAX_COLS + col → the cell's last write. */
  cells: Map<number, CellEntry>;
  styles: StyleCall[];
  merges: MergeCall[];
  columns: ColumnCall[];
  rows: RowCall[];
  frozen?: { rows: number; cols: number };
  /** The description arrays' next indices. */
  next: { data: number; styles: number; columns: number; rows: number; merges: number };
}

/** A validated cell write, recorded only once the whole call validated. */
interface PendingCell {
  key: number;
  entry: CellEntry;
  styles: StyleCall[];
}

const keyOf = (row: number, col: number): number => row * MAX_COLS + col;

// --- addresses ---------------------------------------------------------------------------

/** An A1 cell reference ("B3", "$B$3"). */
function expectCellRef(value: unknown, path: string): { row: number; col: number; text: string } {
  const text = expectString(value, path);
  const ref = parseA1(text);
  if (ref === null) {
    throw new DocumentGenerationError(`${path}: expected a cell reference like "B3", got "${text}"`);
  }
  return { row: ref.row, col: ref.col, text };
}

const COLUMNS = /^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/;
const ROWS = /^\$?([1-9]\d{0,6})(?::\$?([1-9]\d{0,6}))?$/;

/** "B:D" / "B" → the columns, 0-based, ascending; null when not that. */
function columnBand(text: string): { start: number; end: number } | null {
  const match = COLUMNS.exec(text);
  if (!match) {
    return null;
  }
  const a = lettersToCol(match[1]!);
  const b = match[2] === undefined ? a : lettersToCol(match[2]);
  return a === null || b === null ? null : { start: Math.min(a, b), end: Math.max(a, b) };
}

/** "3:5" / "3" → the rows, 0-based, ascending; null when not that. */
function rowBand(text: string): { start: number; end: number } | null {
  const match = ROWS.exec(text);
  if (!match) {
    return null;
  }
  const a = Number(match[1]) - 1;
  const b = match[2] === undefined ? a : Number(match[2]) - 1;
  return a >= MAX_ROWS || b >= MAX_ROWS ? null : { start: Math.min(a, b), end: Math.max(a, b) };
}

/** A range a style call accepts: "A1:D4", "B3", whole columns "B:D" / "B",
 * whole rows "3:5" / "3". */
function expectRange(value: unknown, path: string): GridRange {
  const text = expectString(value, path);
  const cells = parseRange(text);
  if (cells !== null) {
    return cells;
  }
  const columns = columnBand(text);
  if (columns !== null) {
    return { top: 0, bottom: MAX_ROWS - 1, left: columns.start, right: columns.end };
  }
  const rows = rowBand(text);
  if (rows !== null) {
    return { top: rows.start, bottom: rows.end, left: 0, right: MAX_COLS - 1 };
  }
  throw new DocumentGenerationError(`${path}: expected a range like "A1:D4", "B:D" or "3:5", got "${text}"`);
}

// --- dates -------------------------------------------------------------------------------

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/;
const ISO_OFFSET = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}(?::?\d{2})?)$/i;

const isLeapYear = (year: number): boolean => (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;

function daysInMonth(year: number, month: number): number {
  return month === 2 ? (isLeapYear(year) ? 29 : 28) : [4, 6, 9, 11].includes(month) ? 30 : 31;
}

/** A serial and whether it carries a time. */
interface DateValue {
  serial: number;
  hasTime: boolean;
}

/** The serial of calendar fields in `system`, through serial.ts. */
function serialOf(y: number, m: number, d: number, msOfDay: number, system: "1900" | "1904"): DateValue {
  return { serial: dateToSerial(y, m, d, system) + msOfDay / MS_PER_DAY, hasTime: msOfDay > 0 };
}

/** A JS Date, read in UTC. */
function dateFromJs(date: Date, path: string, system: "1900" | "1904"): DateValue {
  if (Number.isNaN(date.getTime())) {
    throw new DocumentGenerationError(`${path}: invalid Date`);
  }
  const year = date.getUTCFullYear();
  if (year < 0 || year > 9999) {
    throw new DocumentGenerationError(`${path}: expected a date in the years 0–9999`);
  }
  const msOfDay = ((date.getUTCHours() * 60 + date.getUTCMinutes()) * 60 + date.getUTCSeconds()) * 1000 + date.getUTCMilliseconds();
  return serialOf(year, date.getUTCMonth() + 1, date.getUTCDate(), msOfDay, system);
}

/** An ISO date string, read by its calendar fields. 1900-02-29 — Excel's
 * fictitious leap day, serial 60 — exists in the 1900 system only. */
function dateFromIso(value: unknown, path: string, system: "1900" | "1904"): DateValue {
  const text = expectString(value, path);
  const match = ISO_DATE.exec(text);
  if (!match) {
    throw new DocumentGenerationError(
      ISO_OFFSET.test(text)
        ? `${path}: "${text}" has a UTC offset — give the calendar date and time to store, without one`
        : `${path}: expected "YYYY-MM-DD[THH:MM[:SS[.sss]]]", got "${text}"`,
    );
  }
  const [y, m, d, hh, mm, ss] = match.slice(1, 7).map((part) => (part === undefined ? 0 : Number(part)));
  const ms = match[7] === undefined ? 0 : Math.round(Number(`0.${match[7]}`) * 1000);
  const lotusLeapDay = system === "1900" && y === 1900 && m === 2 && d === 29;
  if (m! < 1 || m! > 12 || d! < 1 || (d! > daysInMonth(y!, m!) && !lotusLeapDay) || hh! > 23 || mm! > 59 || ss! > 59) {
    throw new DocumentGenerationError(`${path}: "${text}" is not a calendar date and time`);
  }
  return serialOf(y!, m!, d!, ((hh! * 60 + mm!) * 60 + ss!) * 1000 + ms, system);
}

// --- formulas ----------------------------------------------------------------------------

/** Calls `visit` for every node of a formula AST. */
function walk(node: FormulaNode, visit: (node: FormulaNode) => void): void {
  visit(node);
  switch (node.kind) {
    case "function":
      node.args.forEach((arg) => walk(arg, visit));
      break;
    case "unary":
      walk(node.operand, visit);
      break;
    case "binary":
      walk(node.left, visit);
      walk(node.right, visit);
      break;
    case "paren":
      walk(node.inner, visit);
      break;
  }
}

/** A formula validated on entry: SPEC-119 syntax, every function in
 * SPEC-120/127's registry, no defined name, table or external reference
 * (this API builds none). Returns its canonical text and the sheets it
 * names. */
function validateFormula(value: unknown, path: string): { formula: string; sheets: string[] } {
  const text = expectString(value, path);
  const body = text.startsWith("=") ? text.slice(1) : text;
  if (body.trim() === "") {
    throw new DocumentGenerationError(`${path}: expected a formula`);
  }
  let node: FormulaNode;
  try {
    node = parseFormula(body);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      // SPEC-119's message, lower-cased as every message after a path is.
      throw new DocumentGenerationError(`${path}: ${error.message.charAt(0).toLowerCase()}${error.message.slice(1)} at ${error.offset}`);
    }
    throw error;
  }
  const sheets = new Set<string>();
  walk(node, (each) => {
    if (each.kind === "function" && getFunction(each.name) === undefined) {
      throw new DocumentGenerationError(`${path}: unknown function ${each.name}`);
    }
    if (each.kind === "name") {
      throw new DocumentGenerationError(`${path}: unknown name ${each.name} (defined names are not built by this API)`);
    }
    if (each.kind === "structured") {
      throw new DocumentGenerationError(`${path}: ${each.text} is a table reference (tables are not built by this API)`);
    }
    if (each.kind === "external") {
      throw new DocumentGenerationError(`${path}: ${each.text} is an external reference`);
    }
    if (each.kind === "ref") {
      for (const sheet of [each.ref.sheet, each.ref.sheetEnd]) {
        if (sheet !== undefined) {
          sheets.add(sheet);
        }
      }
    }
  });
  return { formula: printFormula(node), sheets: [...sheets] };
}

// --- the plan ----------------------------------------------------------------------------

/** The builder's state (see the module header). */
export class WorkbookPlan {
  readonly title: string | undefined;
  readonly format: WorkbookOutputFormat;
  readonly dateSystem: "1900" | "1904";
  readonly locale: SpreadsheetLocaleTag;
  private readonly sheets: PlanSheet[] = [];
  private current: PlanSheet | undefined;
  private readonly presets = new Map<string, CellStyleSpec>();
  /** A model holding the sheets' names, for validateSheetName. */
  private readonly names: SpreadsheetModel = { ...createBlankWorkbook(), sheets: [] };

  /** `options` is validated with root paths ("format: …"), as the
   * description's own top-level keys are. */
  constructor(options: unknown = {}) {
    const obj = expectPlainObject(options, "options");
    rejectUnknownKeys(obj, OPTION_KEYS, "");
    this.title = obj.title === undefined ? undefined : expectString(obj.title, "title");
    this.format = obj.format === undefined ? "xlsx" : expectEnum(obj.format, WORKBOOK_OUTPUT_FORMATS, "format");
    this.dateSystem = obj.dateSystem === undefined ? "1900" : expectEnum(obj.dateSystem, ["1900", "1904"] as const, "dateSystem");
    this.locale = obj.locale === undefined ? DEFAULT_SPREADSHEET_LOCALE : expectEnum(obj.locale, SPREADSHEET_LOCALE_TAGS, "locale");
  }

  private sheet(): PlanSheet {
    if (this.current === undefined) {
      throw new DocumentGenerationError(NO_SHEET_MESSAGE);
    }
    return this.current;
  }

  /** A named style for later calls to use by name. */
  defineStyle(name: unknown, spec: unknown): void {
    const text = expectString(name, "styles");
    if (text.trim() === "") {
      throw new DocumentGenerationError(`styles: expected a style name`);
    }
    this.presets.set(text, validateStyleSpec(spec, `styles.${text}`));
  }

  /** A style argument: a spec, or a defined style's name. */
  private resolveStyle(value: unknown, path: string): CellStyleSpec {
    if (typeof value === "string") {
      const preset = this.presets.get(value);
      if (preset === undefined) {
        throw new DocumentGenerationError(`${path}: no style called "${value}" — call defineStyle first`);
      }
      return preset;
    }
    return validateStyleSpec(value, path);
  }

  addSheet(name: unknown, options: unknown): void {
    const path = `sheets[${this.sheets.length}]`;
    const text = expectString(name, `${path}.name`);
    const problem = validateSheetName(this.names, text);
    if (problem !== null) {
      throw new DocumentGenerationError(`${path}.name: ${problem}`);
    }
    const obj = expectPlainObject(options ?? {}, path);
    rejectUnknownKeys(obj, SHEET_OPTION_KEYS, path);
    const tabColor = obj.tabColor === undefined ? undefined : expectHexColor(obj.tabColor, `${path}.tabColor`);
    const hidden = obj.hidden === undefined ? false : expectBoolean(obj.hidden, `${path}.hidden`);
    const showGridLines = obj.showGridLines === undefined ? true : expectBoolean(obj.showGridLines, `${path}.showGridLines`);
    const sheet: PlanSheet = {
      name: text,
      path,
      hidden,
      showGridLines,
      cells: new Map(),
      styles: [],
      merges: [],
      columns: [],
      rows: [],
      next: { data: 0, styles: 0, columns: 0, rows: 0, merges: 0 },
    };
    if (tabColor !== undefined) {
      sheet.tabColor = tabColor;
    }
    this.sheets.push(sheet);
    this.names.sheets = [...this.names.sheets, emptySheet(text)];
    this.current = sheet;
  }

  /** Makes an existing sheet (any case) current. */
  selectSheet(name: unknown): void {
    const text = expectString(name, "name");
    const upper = text.toUpperCase();
    const found = this.sheets.find((sheet) => sheet.name.toUpperCase() === upper);
    if (found === undefined) {
      throw new DocumentGenerationError(`no sheet called "${text}" — call addSheet first`);
    }
    this.current = found;
  }

  /** One cell's input validated: the entry and its own style calls. */
  private cellInput(sheet: PlanSheet, row: number, col: number, input: unknown, path: string, parseText: boolean): PendingCell {
    const key = keyOf(row, col);
    const range = { top: row, left: col, bottom: row, right: col };
    const styles: StyleCall[] = [];
    const own = (patch: StylePatch, at: string): void => {
      styles.push({ ranges: [range], patch, path: at, cell: true });
    };
    const text = (value: string, at: string): CellValue => {
      if (value.length > MAX_CELL_TEXT_LENGTH) {
        throw new DocumentGenerationError(`${at}: text longer than 32,767 characters`);
      }
      return value === "" ? null : value;
    };
    const scalar = (value: unknown, at: string): CellValue => {
      if (value === null || typeof value === "boolean") {
        return value;
      }
      if (typeof value === "number") {
        if (!Number.isFinite(value)) {
          throw new DocumentGenerationError(`${at}: expected a finite number`);
        }
        return value === 0 ? 0 : value;
      }
      if (typeof value === "string") {
        return text(value, at);
      }
      throw new DocumentGenerationError(`${at}: expected a string, number, boolean, Date, null or cell object`);
    };
    const entry: CellEntry = { value: null, path };
    if (input instanceof Date) {
      const date = dateFromJs(input, path, this.dateSystem);
      entry.value = date.serial;
      own({ numberFormat: date.hasTime ? DATE_TIME_FORMAT : DATE_FORMAT }, path);
      return { key, entry, styles };
    }
    if (typeof input === "string" && parseText) {
      const inferred = inferCsvField(input, SPREADSHEET_LOCALES[this.locale], this.dateSystem);
      entry.value = typeof inferred.value === "string" ? text(inferred.value, path) : inferred.value;
      if (inferred.numberFormat !== undefined) {
        own(inferred.builtinFormatId === undefined ? { numberFormat: inferred.numberFormat } : { numberFormat: inferred.numberFormat, builtinFormatId: inferred.builtinFormatId }, path);
      }
      return { key, entry, styles };
    }
    if (typeof input !== "object" || input === null || Array.isArray(input)) {
      entry.value = scalar(input, path);
      return { key, entry, styles };
    }
    const obj = expectPlainObject(input, path);
    rejectUnknownKeys(obj, CELL_INPUT_KEYS, path);
    const given = CONTENT_KEYS.filter((name) => obj[name] !== undefined);
    if (given.length > 1) {
      throw new DocumentGenerationError(`${path}: specify at most one of "value", "date", "error" and "formula"`);
    }
    let date: DateValue | undefined;
    if (obj.value !== undefined) {
      entry.value = scalar(obj.value, `${path}.value`);
    } else if (obj.date !== undefined) {
      date = dateFromIso(obj.date, `${path}.date`, this.dateSystem);
      entry.value = date.serial;
    } else if (obj.error !== undefined) {
      entry.value = { error: expectEnum(obj.error, ERROR_CODES, `${path}.error`) };
    } else if (obj.formula !== undefined) {
      const { formula, sheets } = validateFormula(obj.formula, `${path}.formula`);
      entry.formula = formula;
      if (sheets.length > 0) {
        entry.sheets = sheets;
      }
    }
    const style = obj.style === undefined ? undefined : this.resolveStyle(obj.style, `${path}.style`);
    const numberFormat = obj.numberFormat === undefined ? undefined : expectNumberFormat(obj.numberFormat, `${path}.numberFormat`);
    if (date !== undefined && numberFormat === undefined && style?.numberFormat === undefined) {
      own({ numberFormat: date.hasTime ? DATE_TIME_FORMAT : DATE_FORMAT }, `${path}.date`);
    }
    if (style !== undefined) {
      for (const call of styleSpecPatches(style, range)) {
        styles.push({ ...call, path: `${path}.style`, cell: true });
      }
    }
    if (numberFormat !== undefined) {
      own({ numberFormat }, `${path}.numberFormat`);
    }
    return { key, entry, styles };
  }

  private commit(sheet: PlanSheet, pending: readonly PendingCell[]): void {
    for (const { key, entry, styles } of pending) {
      sheet.cells.set(key, entry);
      sheet.styles.push(...styles);
    }
  }

  setCell(ref: unknown, input: unknown): void {
    const sheet = this.sheet();
    const at = expectCellRef(ref, `${sheet.path}.cells`);
    this.commit(sheet, [this.cellInput(sheet, at.row, at.col, input, `${sheet.path}.cells.${at.text}`, false)]);
  }

  /** setRows' and setRow's shared body: a 2-D block from `at`, left to
   * right and top to bottom. */
  setRows(at: unknown, rows: unknown, options: unknown, oneRow = false): void {
    const sheet = this.sheet();
    const path = `${sheet.path}.data[${sheet.next.data}]`;
    const origin = expectCellRef(at, `${path}.at`);
    let parseText = false;
    if (options !== undefined) {
      const obj = expectPlainObject(options, path);
      rejectUnknownKeys(obj, ["parseText"], path);
      parseText = obj.parseText === undefined ? false : expectBoolean(obj.parseText, `${path}.parseText`);
    }
    const block = oneRow ? [expectArray(rows, `${path}.rows[0]`)] : expectArray(rows, `${path}.rows`);
    const pending: PendingCell[] = [];
    block.forEach((values, r) => {
      expectArray(values, `${path}.rows[${r}]`).forEach((input, c) => {
        const cellPath = `${path}.rows[${r}][${c}]`;
        const row = origin.row + r;
        const col = origin.col + c;
        if (row >= MAX_ROWS || col >= MAX_COLS) {
          throw new DocumentGenerationError(`${cellPath}: past the last cell of the sheet (XFD1048576)`);
        }
        pending.push(this.cellInput(sheet, row, col, input, cellPath, parseText));
      });
    });
    sheet.next.data += 1;
    this.commit(sheet, pending);
  }

  /** style(range, spec | name) and numberFormat(range, code), or the
   * description's `styles[k]` holding either or both (style first). */
  style(range: unknown, style: unknown, numberFormat: unknown): void {
    const sheet = this.sheet();
    const path = `${sheet.path}.styles[${sheet.next.styles}]`;
    const target = expectRange(range, `${path}.range`);
    if (style === undefined && numberFormat === undefined) {
      throw new DocumentGenerationError(`${path}: expected "style" or "numberFormat"`);
    }
    const calls: StyleCall[] = [];
    if (style !== undefined) {
      const spec = this.resolveStyle(style, `${path}.style`);
      calls.push(...styleSpecPatches(spec, target).map((call) => ({ ...call, path: `${path}.style` })));
    }
    if (numberFormat !== undefined) {
      calls.push({ ranges: [target], patch: { numberFormat: expectNumberFormat(numberFormat, `${path}.numberFormat`) }, path: `${path}.numberFormat` });
    }
    sheet.next.styles += 1;
    sheet.styles.push(...calls);
  }

  /** Column widths: `{ chars }` by SPEC-117's MDW formula, keeping the
   * characters as the xlsx width; `{ px }` as is. */
  columnWidth(columns: unknown, width: unknown): void {
    const sheet = this.sheet();
    const path = `${sheet.path}.columns[${sheet.next.columns}]`;
    const text = expectString(columns, `${path}.columns`);
    const band = columnBand(text);
    if (band === null) {
      throw new DocumentGenerationError(`${path}.columns: expected a column or columns like "B" or "B:D", got "${text}"`);
    }
    const widthPath = `${path}.width`;
    const obj = expectPlainObject(width, widthPath);
    rejectUnknownKeys(obj, ["chars", "px"], widthPath);
    if ((obj.chars === undefined) === (obj.px === undefined)) {
      throw new DocumentGenerationError(`${widthPath}: specify either "chars" or "px"`);
    }
    let call: ColumnCall;
    if (obj.chars !== undefined) {
      const chars = expectNonNegativeNumber(obj.chars, `${widthPath}.chars`);
      if (chars > MAX_COLUMN_WIDTH_CHARS) {
        throw new DocumentGenerationError(`${widthPath}.chars: expected at most ${MAX_COLUMN_WIDTH_CHARS} characters`);
      }
      call = { ...band, px: xlsxWidthToPx(chars, MDW), chars, path };
    } else {
      const px = expectNonNegativeNumber(obj.px, `${widthPath}.px`);
      if (px > MAX_COLUMN_WIDTH_PX) {
        throw new DocumentGenerationError(`${widthPath}.px: expected at most ${MAX_COLUMN_WIDTH_PX} px (${MAX_COLUMN_WIDTH_CHARS} characters)`);
      }
      call = { ...band, px, path };
    }
    sheet.next.columns += 1;
    sheet.columns.push(call);
  }

  /** Row heights: `rows` 1-based as the user sees them (3, "3" or "3:5"). */
  rowHeight(rows: unknown, height: unknown): void {
    const sheet = this.sheet();
    const path = `${sheet.path}.rows[${sheet.next.rows}]`;
    const band =
      typeof rows === "number" ? (Number.isInteger(rows) && rows >= 1 && rows <= MAX_ROWS ? { start: rows - 1, end: rows - 1 } : null) : rowBand(expectString(rows, `${path}.rows`));
    if (band === null) {
      throw new DocumentGenerationError(`${path}.rows: expected a row or rows like 3, "3" or "3:5"`);
    }
    const heightPath = `${path}.height`;
    const obj = expectPlainObject(height, heightPath);
    rejectUnknownKeys(obj, ["pt", "px"], heightPath);
    if ((obj.pt === undefined) === (obj.px === undefined)) {
      throw new DocumentGenerationError(`${heightPath}: specify either "pt" or "px"`);
    }
    const pt = obj.pt !== undefined ? expectNonNegativeNumber(obj.pt, `${heightPath}.pt`) : expectNonNegativeNumber(obj.px, `${heightPath}.px`) * PT_PER_PX;
    if (pt > MAX_ROW_HEIGHT_PT) {
      throw new DocumentGenerationError(`${heightPath}.${obj.pt !== undefined ? "pt" : "px"}: expected at most ${MAX_ROW_HEIGHT_PT} pt`);
    }
    sheet.next.rows += 1;
    sheet.rows.push({ ...band, pt, path });
  }

  merge(range: unknown): void {
    const sheet = this.sheet();
    const path = `${sheet.path}.merges[${sheet.next.merges}]`;
    const text = expectString(range, path);
    const target = parseRange(text);
    if (target === null || (target.top === target.bottom && target.left === target.right)) {
      throw new DocumentGenerationError(`${path}: expected a range of two or more cells like "A1:D1", got "${text}"`);
    }
    const overlapped = sheet.merges.find((merge) => rangesIntersect(merge.range, target));
    if (overlapped !== undefined) {
      throw new DocumentGenerationError(`${path}: ${formatRange(target)} overlaps the merge ${formatRange(overlapped.range)}`);
    }
    sheet.next.merges += 1;
    sheet.merges.push({ range: target, path });
  }

  freeze(spec: unknown): void {
    const sheet = this.sheet();
    const path = `${sheet.path}.freeze`;
    const obj = expectPlainObject(spec, path);
    rejectUnknownKeys(obj, ["rows", "cols"], path);
    const rows = obj.rows === undefined ? 0 : expectIntInRange(obj.rows, `${path}.rows`, 0, MAX_ROWS - 1);
    const cols = obj.cols === undefined ? 0 : expectIntInRange(obj.cols, `${path}.cols`, 0, MAX_COLS - 1);
    sheet.frozen = { rows, cols };
  }

  // --- materialisation -------------------------------------------------------------------

  /** A fresh model in `format` (see the module header) — the plan is never
   * changed, so one builder serves any number of outputs. */
  materialise(format: WorkbookOutputFormat): SpreadsheetModel {
    const stored = this.sheets.reduce((sum, sheet) => sum + sheet.cells.size, 0);
    if (stored > MAX_WORKBOOK_CELLS) {
      throw new DocumentGenerationError(`sheets: ${WORKBOOK_TOO_LARGE_MESSAGE}`);
    }
    // 1. The blank workbook with the builder's sheets.
    const model = createBlankWorkbook();
    if (this.title !== undefined) {
      model.title = this.title;
    }
    model.format = format;
    model.dateSystem = this.dateSystem;
    if (this.sheets.length > 0) {
      model.sheets = this.sheets.map((sheet) => emptySheet(sheet.name));
    }
    const ids = model.sheets.map((sheet) => sheet.id);
    // 2. Cells into rows.
    const formulas = new Map<CellKey, CellEntry>();
    this.sheets.forEach((sheet, index) => {
      model.sheets[index] = { ...model.sheets[index]!, rows: cellRows(sheet, ids[index]!, formulas) };
    });
    // 3. Style calls, in call order.
    this.sheets.forEach((sheet, index) => applyStyleCalls(model, ids[index]!, sheet.styles));
    // 4. Merges.
    this.sheets.forEach((sheet, index) => {
      for (const merge of sheet.merges) {
        const discarded = discardedBy(model.sheets[index]!, merge.range);
        if (discarded !== null) {
          throw new DocumentGenerationError(`${merge.path}: ${formatRange(merge.range)} would discard the value in ${discarded}`);
        }
        check(mergeCells(model, ids[index]!, [merge.range], "cells"), merge.path);
      }
    });
    // 5. Sizes, panes and the sheet options.
    this.sheets.forEach((sheet, index) => {
      const id = ids[index]!;
      for (const column of sheet.columns) {
        check(resizeColumns(model, id, [{ start: column.start, end: column.end }], column.px, column.chars), column.path);
      }
      for (const row of sheet.rows) {
        check(resizeRows(model, id, [{ start: row.start, end: row.end }], row.pt), row.path);
      }
      if (sheet.frozen !== undefined) {
        check(freezePanes(model, id, sheet.frozen.rows, sheet.frozen.cols), `${sheet.path}.freeze`);
      }
      const next: Sheet = { ...model.sheets[index]! };
      if (sheet.tabColor !== undefined) {
        next.tabColor = sheet.tabColor;
      }
      if (sheet.hidden) {
        next.visibility = "hidden";
      }
      if (!sheet.showGridLines) {
        next.showGridLines = false;
      }
      model.sheets[index] = next;
    });
    // 6. Whole-workbook validation.
    this.checkWorkbook(model, formulas);
    // 7. Every formula's cached value: every formula cell's key, and `full`
    // since every one is new (the pass then builds the dependency graph once
    // rather than updating it cell by cell).
    const result = recalculate(model, [...formulas.keys()], { reason: "edit", full: true, settings: { locale: SPREADSHEET_LOCALES[this.locale] } });
    const cycle = result.circular[0];
    if (cycle !== undefined) {
      throw new DocumentGenerationError(`${formulas.get(cycle)!.path}.formula: circular reference (${this.cellName(model, cycle)})`);
    }
    const stale = result.stale[0];
    if (stale !== undefined) {
      const reason = uncomputableReasonOf(model, stale);
      const why = reason === undefined ? "" : ` (${reason.kind}${"name" in reason ? ` ${reason.name}` : ""})`;
      throw new DocumentGenerationError(`${formulas.get(stale)!.path}.formula: can't be computed${why}`);
    }
    // 8. The structural self-check.
    selfCheck(model);
    return model;
  }

  /** "Sales!B2" for a cell key. */
  private cellName(model: SpreadsheetModel, key: CellKey): string {
    for (const sheet of model.sheets) {
      if (key.startsWith(`${sheet.id}!`)) {
        const [row, col] = key.slice(sheet.id.length + 1).split(",").map(Number);
        return `${quoteSheetName(sheet.name)}!${formatA1({ row: row!, col: col! })}`;
      }
    }
    return key;
  }

  /** Step 6 (see the module header). */
  private checkWorkbook(model: SpreadsheetModel, formulas: ReadonlyMap<CellKey, CellEntry>): void {
    const names = new Set(this.sheets.map((sheet) => sheet.name.toUpperCase()));
    if (this.sheets.length === 0) {
      names.add(model.sheets[0]!.name.toUpperCase());
    }
    for (const entry of formulas.values()) {
      const missing = entry.sheets?.find((name) => !names.has(name.toUpperCase()));
      if (missing !== undefined) {
        throw new DocumentGenerationError(`${entry.path}.formula: refers to a sheet the workbook does not have ("${missing}")`);
      }
    }
    const visible = model.sheets.findIndex((sheet) => sheet.visibility === "visible");
    if (visible < 0) {
      throw new DocumentGenerationError(`${this.sheets.at(-1)!.path}.hidden: ${LAST_VISIBLE_SHEET_MESSAGE}`);
    }
    model.activeSheetIndex = visible;
    const stored = model.sheets.reduce((sum, sheet) => sum + sheet.rows.reduce((cells, row) => cells + row.cells.length, 0), 0);
    if (stored > MAX_WORKBOOK_CELLS) {
      throw new DocumentGenerationError(`sheets: ${WORKBOOK_TOO_LARGE_MESSAGE}`);
    }
  }
}

/** A command's FormatResult, a refusal becoming the error at `path`. */
function check(result: { ok: true } | { ok: false; message: string }, path: string): void {
  if (!result.ok) {
    throw new DocumentGenerationError(`${path}: ${result.message}`);
  }
}

/** Step 2: a sheet's accumulated cells as sorted, sparse rows; each formula
 * cell's entry recorded under its key. */
function cellRows(sheet: PlanSheet, sheetId: string, formulas: Map<CellKey, CellEntry>): Row[] {
  const keys = Float64Array.from(sheet.cells.keys()).toSorted();
  const rows: Row[] = [];
  let row: Row | undefined;
  for (const key of keys) {
    const entry = sheet.cells.get(key)!;
    if (entry.value === null && entry.formula === undefined) {
      continue;
    }
    const index = Math.floor(key / MAX_COLS);
    const col = key - index * MAX_COLS;
    if (row === undefined || row.index !== index) {
      row = { index, cells: [] };
      rows.push(row);
    }
    const cell: Cell = { col, value: entry.value };
    if (entry.formula !== undefined) {
      cell.formula = entry.formula;
      formulas.set(cellKey(sheetId, index, col), entry);
    }
    row.cells.push(cell);
  }
  return rows;
}

/** Step 3 (see the module header): every call through applyStylePatch, a
 * run of cell-own calls batched — one applyStylePatch per distinct patch,
 * the patches in order of first use. A cell may take several patches in one
 * batch (a spec's `all` borders, then its font) as long as it takes them in
 * that order; one taken out of it starts a new batch. */
function applyStyleCalls(model: SpreadsheetModel, sheetId: string, calls: readonly StyleCall[]): void {
  /** The batch's groups, in order of first use, and each one's index by
   * patch JSON. */
  let groups: StyleCall[][] = [];
  let indexOf = new Map<string, number>();
  /** cell → the index of the last group it joined. */
  let joined = new Map<number, number>();
  const flush = (): void => {
    for (const group of groups) {
      check(applyStylePatch(model, sheetId, rectangles(group.map((call) => call.ranges[0]!)), group[0]!.patch), group[0]!.path);
    }
    groups = [];
    indexOf = new Map();
    joined = new Map();
  };
  for (const call of calls) {
    if (!call.cell) {
      flush();
      check(applyStylePatch(model, sheetId, call.ranges, call.patch), call.path);
      continue;
    }
    const range = call.ranges[0]!;
    const cell = keyOf(range.top, range.left);
    const patch = JSON.stringify(call.patch);
    let index = indexOf.get(patch);
    if (index !== undefined && index < (joined.get(cell) ?? -1)) {
      flush();
      index = undefined;
    }
    if (index === undefined) {
      index = groups.length;
      groups.push([]);
      indexOf.set(patch, index);
    }
    groups[index]!.push(call);
    joined.set(cell, index);
  }
  flush();
}

/** Single cells as the fewest rectangles a column-then-row sweep finds:
 * vertical runs per column, joined across adjacent columns holding the same
 * run. */
function rectangles(cells: readonly GridRange[]): GridRange[] {
  const sorted = cells.toSorted((a, b) => a.left - b.left || a.top - b.top);
  const runs: GridRange[] = [];
  for (const cell of sorted) {
    const last = runs.at(-1);
    if (last && last.left === cell.left && last.bottom + 1 >= cell.top) {
      last.bottom = Math.max(last.bottom, cell.top);
    } else {
      runs.push({ top: cell.top, left: cell.left, bottom: cell.top, right: cell.left });
    }
  }
  const open = new Map<string, GridRange>();
  const out: GridRange[] = [];
  for (const run of runs) {
    const key = `${run.top}:${run.bottom}`;
    const rect = open.get(key);
    if (rect && rect.right + 1 === run.left) {
      rect.right = run.left;
    } else {
      const next = { ...run };
      open.set(key, next);
      out.push(next);
    }
  }
  return out;
}

/** The first cell (row by row) a merge of `range` would clear, or null. */
function discardedBy(sheet: Sheet, range: GridRange): string | null {
  for (const row of sheet.rows) {
    if (row.index < range.top || row.index > range.bottom) {
      continue;
    }
    for (const cell of row.cells) {
      if (cell.col < range.left || cell.col > range.right || (row.index === range.top && cell.col === range.left)) {
        continue;
      }
      if (cell.value !== null || cell.formula !== undefined) {
        return `${colToLetters(cell.col)}${row.index + 1}`;
      }
    }
  }
  return null;
}

function fail(why: string): never {
  throw new DocumentGenerationError(`internal: the generated workbook breaks an invariant (${why})`);
}

/** Step 8: SPEC-116's invariants hold — rows ascending and unique, cells
 * ascending and unique, nothing empty stored, no retained source. */
function selfCheck(model: SpreadsheetModel): void {
  if (model.opaque.length > 0) {
    fail("opaque parts");
  }
  for (const sheet of model.sheets) {
    if (sheet.partName !== undefined || sheet.sourceXml !== undefined) {
      fail(`${sheet.name} has a source`);
    }
    let last = -1;
    for (const row of sheet.rows) {
      if (row.index <= last || isEmptyRow(row)) {
        fail(`${sheet.name} row ${row.index + 1}`);
      }
      last = row.index;
      let lastCol = -1;
      for (const cell of row.cells) {
        if (cell.col <= lastCol || (cell.value === null && cell.formula === undefined && cell.styleId === undefined)) {
          fail(`${sheet.name} ${colToLetters(cell.col)}${row.index + 1}`);
        }
        lastCol = cell.col;
      }
    }
  }
}
