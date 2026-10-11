// The internal cell clipboard payload (SPEC-125 "The payload", "Validation")
// — THIS FILE is the normative source for SheetClipboardPayload, its wire
// form and parseClipboardPayload. clipboard.ts builds the payload and the
// flavours around it; paste.ts applies it.
//
// - Wire form: the payload as JSON, encoded as UTF-8, then base64, in the
//   `data-docu-edit-cells` attribute of the text/html flavour's <table>. It
//   rides inside text/html rather than a private MIME type because the async
//   Clipboard API (the cell menu's Paste, Paste special) exposes only
//   text/plain, text/html and images — one standard flavour carries it on
//   every path.
// - The attribute is hostile input, whoever wrote it (SPEC-016's rule,
//   mirroring lib/clipboard.ts's parseFragmentPayload): its size is checked
//   before decoding, UTF-8 is decoded with `fatal: true`, JSON is parsed in a
//   `try`, `kind`/`version` are checked, and every field is READ AND
//   RE-CREATED from a whitelist under the bounds below — nothing is copied by
//   spread, unknown keys are ignored, and only own properties are read (a
//   `__proto__` key is just an unknown key). Any violation returns null and
//   the paste falls through to the HTML-table import; two rules are soft: a
//   formula `parseFormula` rejects is dropped (its cached value kept), and an
//   array range that is not inside the shape anchored at its own cell is
//   dropped.
//
// | Field | Rule |
// | attribute | ≤ 96 MiB of base64, checked before decoding; UTF-8 decoded with `fatal: true` |
// | top level | a plain object; `kind === "docu-edit/cells"`, `version === 1` |
// | `rows`, `cols` | integers ≥ 1, `rows × cols ≤ MAX_COPY_CELLS`; `origin` and the whole rectangle inside the grid |
// | `cells` | ≤ `rows × cols` entries, integer `r < rows`, `c < cols`, strictly increasing row-major |
// | `v` | finite number, string ≤ 32,767 characters, boolean, null, or `{ error }` with a known ErrorCode |
// | `f` | string ≤ 8,192 characters that parseFormula accepts — otherwise dropped and `v` kept |
// | `a` | inside the shape and anchored at its own cell, else dropped |
// | `s` | integer < `styles.length` |
// | `styles` | ≤ 64,000, each rebuilt field by field (see validStyle) |
// | `t` | ≤ 32,767 characters in total, fonts as styles'; runs not spelling `v` are dropped |
// | `merges` | inside the shape, at least two cells, pairwise disjoint |
// | `rowProps`/`colProps` | only with the matching flag; indices inside the shape, ascending; sizes within SPEC-124's dialog bounds |
// | `workbookId`, `sheetName`, `cut` | strings ≤ 255 (`sheetName` 1–31) |
import { base64ToBytes, bytesToBase64 } from "../core/packageParts";
import { THEME_COLOR_SLOTS, type ModelColor, type ModelFont, type ThemeColorRef } from "../core/theme";
import type { CellRef, GridRange } from "./address";
import { MAX_COLUMN_WIDTH_CHARS, MAX_ROW_HEIGHT_PT } from "./bands";
import { parseFormula } from "./formula/parse";
import { MAX_TEXT_LENGTH } from "./formula/values";
import {
  ERROR_CODES,
  MAX_COLS,
  MAX_ROWS,
  type CellAlignment,
  type CellBorderEdge,
  type CellBorders,
  type CellFill,
  type CellFont,
  type CellStyle,
  type CellValue,
  type SheetBorderStyle,
  type TextRun,
} from "./model";
import { BUILTIN_NUMBER_FORMATS } from "./numberFormat/builtin";
import { parseFormatCode } from "./numberFormat/parse";

export interface SheetClipboardPayload {
  kind: "docu-edit/cells";
  version: 1;
  workbookId: string; // source SpreadsheetModel.id
  sheetName: string;
  dateSystem: "1900" | "1904";
  origin: CellRef; // source top-left
  rows: number; // the copied rectangle after used-range clipping (SPEC-128: its visible rows, packed, when a copy leaves filter-hidden rows out)
  cols: number;
  wholeRows?: true;
  wholeCols?: true;
  cells: ClipboardCell[]; // sparse, row-major ascending, unique (r, c)
  styles: CellStyle[]; // distinct styles referenced; no sourceXf/sourceStyleName
  merges: GridRange[]; // relative to origin
  rowProps?: ClipboardRowProps[]; // wholeRows
  colProps?: ClipboardColProps[]; // wholeCols
  cut?: string; // the cut token
}

export interface ClipboardCell {
  r: number; // relative to origin
  c: number;
  v: CellValue; // the value, or a formula's cached value
  f?: string; // canonical formula
  a?: GridRange; // relative arrayRange — only when the whole array was copied
  s?: number; // index into styles
  t?: TextRun[]; // rich text
}

export interface ClipboardRowProps {
  r: number;
  heightPt?: number;
  customHeight?: true;
  hidden?: true;
  s?: number;
}

export interface ClipboardColProps {
  c: number;
  widthPx?: number;
  hidden?: true;
  s?: number;
}

export const MAX_COPY_CELLS = 1_000_000;
export const PAYLOAD_KIND = "docu-edit/cells";
export const PAYLOAD_ATTRIBUTE = "data-docu-edit-cells";
/** 96 MiB of base64 (≈ 64 MiB of JSON), checked before decoding. */
export const MAX_PAYLOAD_ATTRIBUTE_LENGTH = 96 * 1024 * 1024;
export const MAX_CLIPBOARD_STYLES = 64_000;
/** Excel's limit on a formula's length, "=" excluded (SPEC-122's). */
export const MAX_CLIPBOARD_FORMULA_LENGTH = 8192;
/** The widest column SPEC-124's width dialog can make (255 characters),
 * at the widest maximum digit width a default font reaches. */
export const MAX_CLIPBOARD_COLUMN_WIDTH_PX = MAX_COLUMN_WIDTH_CHARS * 16 + 5;
const MAX_NAME_LENGTH = 255;
const MAX_SHEET_NAME_LENGTH = 31;

// --- the wire form ---------------------------------------------------------------------

/** The payload's attribute text: JSON → UTF-8 → base64. */
export function encodePayload(payload: SheetClipboardPayload): string {
  return bytesToBase64(new TextEncoder().encode(JSON.stringify(payload)));
}

// --- validation ------------------------------------------------------------------------

class Invalid extends Error {}

function fail(): never {
  throw new Invalid("invalid clipboard payload");
}

type Data = Record<string, unknown>;

function isRecord(value: unknown): value is Data {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** An own property only — never one inherited through a prototype. */
function own(data: Data, key: string): unknown {
  return Object.prototype.hasOwnProperty.call(data, key) ? data[key] : undefined;
}

function record(value: unknown): Data {
  return isRecord(value) ? value : fail();
}

function int(value: unknown, min: number, max: number): number {
  return typeof value === "number" && Number.isInteger(value) && value >= min && value <= max ? value : fail();
}

function num(value: unknown, min: number, max: number): number {
  return typeof value === "number" && Number.isFinite(value) && value >= min && value <= max ? value : fail();
}

function text(value: unknown, max: number, min = 0): string {
  return typeof value === "string" && value.length >= min && value.length <= max ? value : fail();
}

/** `true`, or absent (undefined); anything else fails. */
function flag(value: unknown): true | undefined {
  if (value === undefined) {
    return undefined;
  }
  return value === true ? true : fail();
}

function oneOf<T extends string>(value: unknown, allowed: readonly T[]): T {
  return typeof value === "string" && (allowed as readonly string[]).includes(value) ? (value as T) : fail();
}

function list(value: unknown, max: number): unknown[] {
  return Array.isArray(value) && value.length <= max ? value : fail();
}

/** Non-empty and free of C0 and C1 control characters. */
export function isPrintable(candidate: string): boolean {
  if (candidate.length === 0) {
    return false;
  }
  for (let i = 0; i < candidate.length; i += 1) {
    const code = candidate.charCodeAt(i);
    if (code < 0x20 || (code >= 0x7f && code <= 0x9f)) {
      return false;
    }
  }
  return true;
}

const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;

const ERROR_CODE_SET: ReadonlySet<string> = new Set(ERROR_CODES);

function validValue(value: unknown): CellValue {
  if (value === null || typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : fail();
  }
  if (typeof value === "string") {
    return text(value, MAX_TEXT_LENGTH);
  }
  const error = own(record(value), "error");
  return typeof error === "string" && ERROR_CODE_SET.has(error) ? { error: error as (typeof ERROR_CODES)[number] } : fail();
}

function validColor(value: unknown): ModelColor {
  if (typeof value === "string") {
    return HEX_COLOR.test(value) ? value.toLowerCase() : fail();
  }
  const raw = record(value);
  const color: ThemeColorRef = { theme: oneOf(own(raw, "theme"), THEME_COLOR_SLOTS) };
  for (const key of ["lumMod", "lumOff", "shade", "tint"] as const) {
    const amount = own(raw, key);
    if (amount !== undefined) {
      color[key] = num(amount, 0, 1);
    }
  }
  return color;
}

function validFamily(value: unknown): ModelFont {
  if (typeof value === "string") {
    return value.length <= MAX_NAME_LENGTH && isPrintable(value) ? value : fail();
  }
  return { theme: oneOf(own(record(value), "theme"), ["major", "minor"] as const) };
}

const UNDERLINES = ["single", "double", "singleAccounting", "doubleAccounting"] as const;

function validFont(value: unknown): CellFont {
  const raw = record(value);
  const font: CellFont = {};
  const family = own(raw, "family");
  if (family !== undefined) {
    font.family = validFamily(family);
  }
  const size = own(raw, "sizePt");
  if (size !== undefined) {
    font.sizePt = num(size, 1, 409);
  }
  for (const key of ["bold", "italic", "strike"] as const) {
    const on = flag(own(raw, key));
    if (on) {
      font[key] = on;
    }
  }
  const underline = own(raw, "underline");
  if (underline !== undefined) {
    font.underline = oneOf(underline, UNDERLINES);
  }
  const color = own(raw, "color");
  if (color !== undefined) {
    font.color = validColor(color);
  }
  const vertAlign = own(raw, "vertAlign");
  if (vertAlign !== undefined) {
    font.vertAlign = oneOf(vertAlign, ["superscript", "subscript"] as const);
  }
  return font;
}

const PATTERNS = [
  "darkGray",
  "mediumGray",
  "lightGray",
  "gray125",
  "gray0625",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
] as const;

const MAX_GRADIENT_STOPS = 32;

function validFill(value: unknown): CellFill {
  const raw = record(value);
  const kind = oneOf(own(raw, "kind"), ["solid", "pattern", "gradient"] as const);
  if (kind === "solid") {
    return { kind, color: validColor(own(raw, "color")) };
  }
  if (kind === "pattern") {
    const fill: CellFill = { kind, pattern: oneOf(own(raw, "pattern"), PATTERNS) };
    const fg = own(raw, "fg");
    const bg = own(raw, "bg");
    if (fg !== undefined) {
      fill.fg = validColor(fg);
    }
    if (bg !== undefined) {
      fill.bg = validColor(bg);
    }
    return fill;
  }
  const stops = list(own(raw, "stops"), MAX_GRADIENT_STOPS).map((stop) => {
    const entry = record(stop);
    return { position: num(own(entry, "position"), 0, 1), color: validColor(own(entry, "color")) };
  });
  const fill: Extract<CellFill, { kind: "gradient" }> = { kind, type: oneOf(own(raw, "type"), ["linear", "path"] as const), stops };
  const degree = own(raw, "degree");
  if (degree !== undefined) {
    fill.degree = num(degree, -360, 360);
  }
  for (const key of ["left", "right", "top", "bottom"] as const) {
    const at = own(raw, key);
    if (at !== undefined) {
      fill[key] = num(at, 0, 1);
    }
  }
  return fill;
}

const BORDER_STYLES: readonly SheetBorderStyle[] = [
  "thin",
  "medium",
  "thick",
  "dashed",
  "dotted",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
];

function validEdge(value: unknown): CellBorderEdge {
  const raw = record(value);
  const edge: CellBorderEdge = { style: oneOf(own(raw, "style"), BORDER_STYLES) };
  const color = own(raw, "color");
  if (color !== undefined) {
    edge.color = validColor(color);
  }
  return edge;
}

function validBorders(value: unknown): CellBorders {
  const raw = record(value);
  const borders: CellBorders = {};
  for (const key of ["top", "right", "bottom", "left"] as const) {
    const edge = own(raw, key);
    if (edge !== undefined) {
      borders[key] = validEdge(edge);
    }
  }
  const diagonal = own(raw, "diagonal");
  if (diagonal !== undefined) {
    const raw2 = record(diagonal);
    const edge: CellBorderEdge & { up?: true; down?: true } = validEdge(raw2);
    const up = flag(own(raw2, "up"));
    const down = flag(own(raw2, "down"));
    if (up) {
      edge.up = up;
    }
    if (down) {
      edge.down = down;
    }
    borders.diagonal = edge;
  }
  return borders;
}

const HORIZONTAL = ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"] as const;
const VERTICAL = ["top", "center", "bottom", "justify", "distributed"] as const;

function validAlignment(value: unknown): CellAlignment {
  const raw = record(value);
  const alignment: CellAlignment = {};
  const horizontal = own(raw, "horizontal");
  if (horizontal !== undefined) {
    alignment.horizontal = oneOf(horizontal, HORIZONTAL);
  }
  const vertical = own(raw, "vertical");
  if (vertical !== undefined) {
    alignment.vertical = oneOf(vertical, VERTICAL);
  }
  for (const key of ["wrap", "shrinkToFit"] as const) {
    const on = flag(own(raw, key));
    if (on) {
      alignment[key] = on;
    }
  }
  const indent = own(raw, "indent");
  if (indent !== undefined) {
    alignment.indent = int(indent, 0, 250);
  }
  const rotation = own(raw, "textRotation");
  if (rotation !== undefined) {
    const degrees = int(rotation, 0, 255);
    alignment.textRotation = degrees <= 180 || degrees === 255 ? degrees : fail();
  }
  return alignment;
}

function validNumberFormat(value: unknown): string {
  const code = text(value, MAX_NAME_LENGTH, 1);
  try {
    parseFormatCode(code);
  } catch {
    fail();
  }
  return code;
}

/** A style rebuilt field by field; never `sourceXf`/`sourceStyleName`. */
function validStyle(value: unknown): CellStyle {
  const raw = record(value);
  const style: CellStyle = {};
  const font = own(raw, "font");
  if (font !== undefined) {
    style.font = validFont(font);
  }
  const fill = own(raw, "fill");
  if (fill !== undefined) {
    style.fill = validFill(fill);
  }
  const borders = own(raw, "borders");
  if (borders !== undefined) {
    style.borders = validBorders(borders);
  }
  const alignment = own(raw, "alignment");
  if (alignment !== undefined) {
    style.alignment = validAlignment(alignment);
  }
  const numberFormat = own(raw, "numberFormat");
  if (numberFormat !== undefined) {
    style.numberFormat = validNumberFormat(numberFormat);
  }
  const builtin = own(raw, "builtinFormatId");
  if (builtin !== undefined) {
    const id = int(builtin, 0, 163);
    style.builtinFormatId = BUILTIN_NUMBER_FORMATS[id] !== undefined ? id : fail();
  }
  const protection = own(raw, "protection");
  if (protection !== undefined) {
    const rawProtection = record(protection);
    const locked = own(rawProtection, "locked");
    const hidden = flag(own(rawProtection, "hidden"));
    style.protection = {};
    if (locked !== undefined) {
      style.protection.locked = locked === false ? false : fail();
    }
    if (hidden) {
      style.protection.hidden = hidden;
    }
  }
  const quotePrefix = flag(own(raw, "quotePrefix"));
  if (quotePrefix) {
    style.quotePrefix = quotePrefix;
  }
  const named = own(raw, "namedStyle");
  if (named !== undefined) {
    style.namedStyle = text(named, MAX_NAME_LENGTH, 1);
  }
  return style;
}

function validRuns(value: unknown): TextRun[] {
  let total = 0;
  return list(value, MAX_TEXT_LENGTH).map((entry) => {
    const raw = record(entry);
    const run: TextRun = { text: text(own(raw, "text"), MAX_TEXT_LENGTH) };
    total += run.text.length;
    if (total > MAX_TEXT_LENGTH) {
      fail();
    }
    const font = own(raw, "font");
    if (font !== undefined) {
      run.font = validFont(font);
    }
    return run;
  });
}

/** A relative range inside the rows × cols shape. */
function validRange(value: unknown, rows: number, cols: number): GridRange {
  const raw = record(value);
  const top = int(own(raw, "top"), 0, rows - 1);
  const left = int(own(raw, "left"), 0, cols - 1);
  const bottom = int(own(raw, "bottom"), top, rows - 1);
  const right = int(own(raw, "right"), left, cols - 1);
  return { top, left, bottom, right };
}

/** Formula texts parseFormula accepted or rejected within one parse. */
type FormulaVerdicts = Map<string, boolean>;

function formulaParses(formula: string, verdicts: FormulaVerdicts): boolean {
  let ok = verdicts.get(formula);
  if (ok === undefined) {
    try {
      parseFormula(formula);
      ok = true;
    } catch {
      ok = false;
    }
    verdicts.set(formula, ok);
  }
  return ok;
}

function validCell(value: unknown, rows: number, cols: number, styleCount: number, verdicts: FormulaVerdicts): ClipboardCell {
  const raw = record(value);
  const r = int(own(raw, "r"), 0, rows - 1);
  const c = int(own(raw, "c"), 0, cols - 1);
  if (!Object.prototype.hasOwnProperty.call(raw, "v")) {
    fail();
  }
  const cell: ClipboardCell = { r, c, v: validValue(own(raw, "v")) };
  const formula = own(raw, "f");
  // Soft: a formula that is not one is dropped, its cached value kept.
  if (typeof formula === "string" && formula.length > 0 && formula.length <= MAX_CLIPBOARD_FORMULA_LENGTH && formulaParses(formula, verdicts)) {
    cell.f = formula;
  }
  const array = own(raw, "a");
  if (array !== undefined && cell.f !== undefined) {
    // Soft: an array range outside the shape, or not anchored here, is dropped.
    try {
      const range = validRange(array, rows, cols);
      if (range.top === r && range.left === c) {
        cell.a = range;
      }
    } catch (error) {
      if (!(error instanceof Invalid)) {
        throw error;
      }
    }
  }
  const style = own(raw, "s");
  if (style !== undefined) {
    cell.s = int(style, 0, styleCount - 1);
  }
  const runs = own(raw, "t");
  if (runs !== undefined) {
    const t = validRuns(runs);
    // Soft: runs that do not spell the value are no rich text of it.
    if (typeof cell.v === "string" && t.map((run) => run.text).join("") === cell.v) {
      cell.t = t;
    }
  }
  return cell;
}

/** Merges inside the shape, at least two cells, pairwise disjoint — checked
 * on an occupancy grid of the shape (at most MAX_COPY_CELLS bytes). */
function validMerges(value: unknown, rows: number, cols: number): GridRange[] {
  const entries = list(value, Math.floor((rows * cols) / 2));
  if (entries.length === 0) {
    return [];
  }
  const taken = new Uint8Array(rows * cols);
  return entries.map((entry) => {
    const range = validRange(entry, rows, cols);
    if (range.top === range.bottom && range.left === range.right) {
      fail();
    }
    for (let r = range.top; r <= range.bottom; r += 1) {
      for (let c = range.left; c <= range.right; c += 1) {
        const at = r * cols + c;
        if (taken[at]) {
          fail();
        }
        taken[at] = 1;
      }
    }
    return range;
  });
}

function validRowProps(value: unknown, rows: number, styleCount: number): ClipboardRowProps[] {
  let previous = -1;
  return list(value, rows).map((entry) => {
    const raw = record(entry);
    const props: ClipboardRowProps = { r: int(own(raw, "r"), previous + 1, rows - 1) };
    previous = props.r;
    const height = own(raw, "heightPt");
    if (height !== undefined) {
      props.heightPt = num(height, 0, MAX_ROW_HEIGHT_PT);
    }
    const custom = flag(own(raw, "customHeight"));
    const hidden = flag(own(raw, "hidden"));
    if (custom) {
      props.customHeight = custom;
    }
    if (hidden) {
      props.hidden = hidden;
    }
    const style = own(raw, "s");
    if (style !== undefined) {
      props.s = int(style, 0, styleCount - 1);
    }
    return props;
  });
}

function validColProps(value: unknown, cols: number, styleCount: number): ClipboardColProps[] {
  let previous = -1;
  return list(value, cols).map((entry) => {
    const raw = record(entry);
    const props: ClipboardColProps = { c: int(own(raw, "c"), previous + 1, cols - 1) };
    previous = props.c;
    const width = own(raw, "widthPx");
    if (width !== undefined) {
      props.widthPx = num(width, 0, MAX_CLIPBOARD_COLUMN_WIDTH_PX);
    }
    const hidden = flag(own(raw, "hidden"));
    if (hidden) {
      props.hidden = hidden;
    }
    const style = own(raw, "s");
    if (style !== undefined) {
      props.s = int(style, 0, styleCount - 1);
    }
    return props;
  });
}

function decodeAttribute(attribute: string): unknown {
  if (attribute.length === 0 || attribute.length > MAX_PAYLOAD_ATTRIBUTE_LENGTH) {
    fail();
  }
  let json: string;
  try {
    json = new TextDecoder("utf-8", { fatal: true }).decode(base64ToBytes(attribute));
  } catch {
    fail();
  }
  try {
    return JSON.parse(json);
  } catch {
    fail();
  }
}

function validPayload(attribute: string): SheetClipboardPayload {
  const raw = record(decodeAttribute(attribute));
  if (own(raw, "kind") !== PAYLOAD_KIND || own(raw, "version") !== 1) {
    fail();
  }
  const originRaw = record(own(raw, "origin"));
  const origin: CellRef = { row: int(own(originRaw, "row"), 0, MAX_ROWS - 1), col: int(own(originRaw, "col"), 0, MAX_COLS - 1) };
  const rows = int(own(raw, "rows"), 1, MAX_ROWS - origin.row);
  const cols = int(own(raw, "cols"), 1, MAX_COLS - origin.col);
  if (rows * cols > MAX_COPY_CELLS) {
    fail();
  }
  const styles = list(own(raw, "styles"), MAX_CLIPBOARD_STYLES).map(validStyle);
  const verdicts: FormulaVerdicts = new Map();
  let previous = -1;
  const cells = list(own(raw, "cells"), rows * cols).map((entry) => {
    const cell = validCell(entry, rows, cols, styles.length, verdicts);
    const at = cell.r * cols + cell.c;
    if (at <= previous) {
      fail();
    }
    previous = at;
    return cell;
  });
  const payload: SheetClipboardPayload = {
    kind: PAYLOAD_KIND,
    version: 1,
    workbookId: text(own(raw, "workbookId"), MAX_NAME_LENGTH),
    sheetName: text(own(raw, "sheetName"), MAX_SHEET_NAME_LENGTH, 1),
    dateSystem: oneOf(own(raw, "dateSystem"), ["1900", "1904"] as const),
    origin,
    rows,
    cols,
    cells,
    styles,
    merges: validMerges(own(raw, "merges"), rows, cols),
  };
  const wholeRows = flag(own(raw, "wholeRows"));
  const wholeCols = flag(own(raw, "wholeCols"));
  if (wholeRows) {
    payload.wholeRows = wholeRows;
  }
  if (wholeCols) {
    payload.wholeCols = wholeCols;
  }
  const rowProps = own(raw, "rowProps");
  if (rowProps !== undefined) {
    payload.rowProps = wholeRows ? validRowProps(rowProps, rows, styles.length) : fail();
  }
  const colProps = own(raw, "colProps");
  if (colProps !== undefined) {
    payload.colProps = wholeCols ? validColProps(colProps, cols, styles.length) : fail();
  }
  const cut = own(raw, "cut");
  if (cut !== undefined) {
    payload.cut = text(cut, MAX_NAME_LENGTH, 1);
  }
  return payload;
}

/** The payload an attribute carries, rebuilt from a whitelist under the
 * bounds of "Validation" — or null on any doubt (see the module header). */
export function parseClipboardPayload(attribute: string): SheetClipboardPayload | null {
  try {
    return validPayload(attribute);
  } catch (error) {
    if (error instanceof Invalid) {
      return null;
    }
    // Anything else thrown while validating is a doubt too: a payload is
    // never the reason a paste crashes.
    return null;
  }
}
