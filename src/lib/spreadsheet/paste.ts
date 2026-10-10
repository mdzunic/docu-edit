// Pasting cells (SPEC-125 "Paste geometry", "Paste semantics", "TSV in") —
// THIS FILE is the normative source for how a clipboard block lands on a
// sheet: the internal payload (clipboardPayload.ts), an external table
// (clipboardImport.ts) or plain text, each first made a PasteBlock.
//
// Geometry:
// | Clipboard | Destination selection | Result |
// | N×M | one cell (or one merge) | pasted at it |
// | 1×1 | any single range | tiled over it |
// | N×M | a range whose height is a multiple of N and width of M | tiled |
// | N×M | any other single range | once, at its top-left |
// | whole rows/columns | any | onto whole rows/columns starting at the active cell's row/column; Transpose disabled |
// | any | several ranges | refused |
// The destination rectangle is the union of the tiles; it is checked against
// the grid edge, MAX_COPY_CELLS, merges (a merge entirely inside is
// replaced, one crossing its edge refuses — except a single cell pasted into
// a merge's anchor, which writes the anchor and keeps the merge), arrays
// (PARTIAL_ARRAY_MESSAGE) and protection (SPEC-122's message) before
// anything changes.
//
// Semantics:
// | Option | Values / formulas | Styles | Merges |
// | All | replaced; formulas shiftFormula'd by (destination − origin) | replaced (blank source → default) | replaced |
// | Formulas | formulas shifted, constants as values | kept | kept |
// | Values | cached values; formulas and runs dropped | kept | kept |
// | Formats | kept | replaced (+ row/column properties for whole rows/columns) | replaced |
// | Values and number formats | as Values | only numberFormat, merged with patchStyle | kept |
// | Skip blanks | a source position with no value and no formula leaves its destination untouched under every option |
// | Transpose | rows ↔ columns; a reference relative on both axes has its offset from its host transposed, absolute and mixed ones shift by the host's displacement; merges transposed |
// - A blank source position clears its destination's value (and, where
//   styles are replaced, its style — to the default, or for whole rows and
//   columns to the row's or column's); preserved content the model does not
//   model (notes) stays. Positions a kept merge covers, other than its
//   anchor, are not written.
// - Same-workbook pastes keep style ids by identity (the caller's in-memory
//   `styleIds`, checked against the styles); other workbooks' styles go
//   through internStyle. Date-formatted values convert between the 1900 and
//   1904 date systems (±1462 days). A formula from another workbook naming a
//   sheet this one lacks (or another workbook) is pasted as its cached value
//   and counted.
// - Plain text (TSV) pastes each field through SPEC-122's parseInput as an
//   entry (cells.ts writeEntries): destination styles kept but for the
//   number format parseInput implies (over General); an empty field clears
//   its destination's value. Paste (All) replaces the destination's merges
//   — text brings none, so a merge inside is gone and every field shows.
//   Under Values a formula field pastes as its text (as typed after an
//   apostrophe), and text with no field but empty ones (a lone line break)
//   is nothing to paste.
// - Google Sheets' R1C1 formulas are converted at each destination cell
//   (formula/r1c1.ts; transposed with the block); an unconvertible one, or
//   one naming a sheet this workbook lacks (counted), pastes its value.
// The pasted sheet is marked dirty; the caller wraps the paste in ONE undo
// step and refits its rows.
import { sanitizeXmlText } from "../core/xmlText";
import type { CellRef, GridRange } from "./address";
import { entryFromInput, firstLockedCell, PROTECTED_CELL_MESSAGE, writeEntries, type CellEntry } from "./cells";
import {
  MULTIPLE_SELECTIONS_MESSAGE,
  PASTE_MERGE_MESSAGE,
  PASTE_PAST_EDGE_MESSAGE,
  parseTsv,
  TOO_MANY_TO_CHANGE_MESSAGE,
} from "./clipboard";
import { MAX_CLIPBOARD_FORMULA_LENGTH, MAX_COPY_CELLS, type ClipboardCell, type ClipboardColProps, type ClipboardRowProps, type SheetClipboardPayload } from "./clipboardPayload";
import { patchStyle } from "./formatting";
import { parseFormula } from "./formula/parse";
import { r1c1ToA1 } from "./formula/r1c1";
import { rewriteReferences, shiftFormula, type RefPart, type ReferenceNode } from "./formula/references";
import { parseSheetPrefix, sheetPrefixOf, tokenize } from "./formula/tokens";
import { MAX_TEXT_LENGTH } from "./formula/values";
import { rangesIntersect } from "./merges";
import { MAX_COLS, MAX_ROWS, type CellStyle, type CellValue, type ColumnSpan, type Sheet, type SpreadsheetModel } from "./model";
import { BUILTIN_NUMBER_FORMATS } from "./numberFormat/builtin";
import { parseInput } from "./numberFormat/input";
import type { SpreadsheetLocale } from "./numberFormat/locale";
import { dateKindOf, DATE_SYSTEM_OFFSET } from "./series";
import { cellWith, contains, inheritedStyleId, SheetEdit, type CellChange } from "./sheetEdit";
import { joinSpans, NOT_A_GRID_MESSAGE, PARTIAL_ARRAY_MESSAGE } from "./structure";
import { cellStyle, internStyle, styleKey } from "./styles";

export type PasteSource = "internal" | "excel" | "sheets" | "libreoffice" | "html" | "text";

export interface PasteCell extends ClipboardCell {
  /** Google Sheets: the formula in R1C1, converted at each destination. */
  r1c1?: string;
  /** Plain text: the entry SPEC-122's parseInput made of the field. */
  entry?: CellEntry;
  /** Plain text, a formula field: the field as text — as typed after an
   * apostrophe — which is what Values pastes. */
  asText?: CellEntry;
}

/** Anything about to be pasted: the payload's shape plus its source. */
export interface PasteBlock {
  source: PasteSource;
  workbookId: string;
  sheetName: string;
  dateSystem: "1900" | "1904";
  origin: CellRef;
  rows: number;
  cols: number;
  wholeRows?: true;
  wholeCols?: true;
  cells: PasteCell[];
  styles: CellStyle[];
  merges: GridRange[];
  rowProps?: ClipboardRowProps[];
  colProps?: ClipboardColProps[];
  cut?: string;
}

export type PasteWhat = "all" | "formulas" | "values" | "formats" | "values-number-formats";

export interface PasteOptions {
  what: PasteWhat;
  skipBlanks?: boolean;
  transpose?: boolean;
}

/** What Paste special offers for a block (options its source cannot supply
 * are disabled): All and Values always. */
export interface PasteCapabilities {
  formulas: boolean;
  formats: boolean;
  numberFormats: boolean;
  transpose: boolean;
}

/** The destination: the sheet's selection. */
export interface PasteTarget {
  ranges: readonly GridRange[];
  active: CellRef;
}

export type PasteResult = { ok: true; range: GridRange; formulasAsValues: number } | { ok: false; message: string };

// --- blocks ----------------------------------------------------------------------------------

export function blockFromPayload(payload: SheetClipboardPayload): PasteBlock {
  const { kind: _kind, version: _version, ...rest } = payload;
  return { source: "internal", ...rest };
}

export interface ParsedField {
  value: CellValue;
  formula?: string;
  numberFormat?: string;
  builtinFormatId?: number;
  quotePrefix?: true;
}

/** How pasted text is read: parseInput's locale and date system, and the
 * view's canonical form of a formula body (SPEC-122's commit — cellEntry.ts
 * canonicalFormula; null: no formula). Without one a formula body SPEC-119
 * parses is kept as written. */
export interface TextContext {
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
  canonical?: (body: string) => string | null;
}

/** A body kept as written when parseFormula accepts it. */
function parsesAsWritten(body: string): string | null {
  if (body.length > MAX_CLIPBOARD_FORMULA_LENGTH) {
    return null;
  }
  try {
    parseFormula(body);
    return body;
  } catch {
    return null;
  }
}

/** A field of pasted text as SPEC-122 reads typed text (parseInput in the
 * locale and date system), a formula in the context's canonical form; text
 * that starts with "=" but is no formula SPEC-119 parses stays text. */
export function parsedField(text: string, ctx: TextContext): ParsedField {
  const clean = sanitizeXmlText(text).slice(0, MAX_TEXT_LENGTH);
  const parsed = parseInput(clean, ctx.locale, ctx.dateSystem);
  if (parsed.formula === undefined) {
    const field: ParsedField = { value: parsed.value };
    if (parsed.numberFormat !== undefined) {
      field.numberFormat = parsed.numberFormat;
    }
    if (parsed.builtinFormatId !== undefined) {
      field.builtinFormatId = parsed.builtinFormatId;
    }
    if (parsed.quotePrefix) {
      field.quotePrefix = true;
    }
    return field;
  }
  const formula = (ctx.canonical ?? parsesAsWritten)(parsed.formula);
  // Not a formula after all: the text as typed.
  return formula === null ? { value: clean } : { value: null, formula };
}

/** Plain text as a block of entries (see the module header); null for no
 * text — or none but empty fields, as a lone line break is: that never
 * clears the selection — a refusal past MAX_COPY_CELLS. */
export function blockFromText(text: string, ctx: TextContext): PasteBlock | string | null {
  const rows = parseTsv(text);
  const cols = rows.reduce((most, fields) => Math.max(most, fields.length), 0);
  if (rows.length === 0 || cols === 0) {
    return null;
  }
  if (rows.length * cols > MAX_COPY_CELLS || rows.length > MAX_ROWS || cols > MAX_COLS) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  const cells: PasteCell[] = [];
  rows.forEach((fields, r) =>
    fields.forEach((field, c) => {
      if (field === "") {
        return;
      }
      const parsed = parsedField(field, ctx);
      const cell: PasteCell = { r, c, v: parsed.value, entry: entryFromInput(parsed) };
      if (parsed.formula !== undefined) {
        cell.f = parsed.formula;
        cell.asText = entryFromInput(parsedField(`'${field}`, ctx));
      }
      cells.push(cell);
    }),
  );
  if (cells.length === 0) {
    return null;
  }
  return { source: "text", workbookId: "", sheetName: "", dateSystem: ctx.dateSystem, origin: { row: 0, col: 0 }, rows: rows.length, cols, cells, styles: [], merges: [] };
}

/** What Paste special offers for a block's source. */
export function pasteCapabilities(block: PasteBlock): PasteCapabilities {
  const transpose = !block.wholeRows && !block.wholeCols;
  switch (block.source) {
    case "internal":
    case "sheets":
      return { formulas: true, formats: true, numberFormats: true, transpose };
    case "text":
      return { formulas: false, formats: false, numberFormats: true, transpose };
    default:
      return { formulas: false, formats: true, numberFormats: true, transpose };
  }
}

// --- geometry --------------------------------------------------------------------------------

/** Where a block lands: the destination rectangle and its tiles (`down` ×
 * `across` copies of the `rows` × `cols` shape from `top`/`left`). */
export interface PastePlan {
  dest: GridRange;
  top: number;
  left: number;
  rows: number;
  cols: number;
  down: number;
  across: number;
  whole?: "rows" | "cols" | "sheet";
  transposed: boolean;
  /** A single cell pasted into a merge's anchor: the merge is kept. */
  intoMerge?: GridRange;
}

/** The written rectangle of a plan (its dest, but for whole rows/columns
 * only the block's own extent). */
export function writtenRange(plan: PastePlan): GridRange {
  return { top: plan.top, left: plan.left, bottom: plan.top + plan.down * plan.rows - 1, right: plan.left + plan.across * plan.cols - 1 };
}

/** The plan for pasting `block` onto the selection (see "Geometry"), or the
 * refusal for a multi-range selection, the grid edge or the cell cap. */
export function planPaste(sheet: Sheet, target: PasteTarget, block: PasteBlock, transpose = false): PastePlan | string {
  if (target.ranges.length !== 1) {
    return MULTIPLE_SELECTIONS_MESSAGE;
  }
  const whole = block.wholeRows && block.wholeCols ? "sheet" : block.wholeRows ? "rows" : block.wholeCols ? "cols" : undefined;
  const transposed = transpose && whole === undefined;
  const rows = transposed ? block.cols : block.rows;
  const cols = transposed ? block.rows : block.cols;
  let plan: PastePlan;
  if (whole === "sheet") {
    plan = { dest: { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }, top: block.origin.row, left: block.origin.col, rows, cols, down: 1, across: 1, whole, transposed };
  } else if (whole === "rows") {
    const top = target.active.row;
    plan = { dest: { top, left: 0, bottom: top + rows - 1, right: MAX_COLS - 1 }, top, left: block.origin.col, rows, cols, down: 1, across: 1, whole, transposed };
  } else if (whole === "cols") {
    const left = target.active.col;
    plan = { dest: { top: 0, left, bottom: MAX_ROWS - 1, right: left + cols - 1 }, top: block.origin.row, left, rows, cols, down: 1, across: 1, whole, transposed };
  } else {
    const selection = target.ranges[0]!;
    const height = selection.bottom - selection.top + 1;
    const width = selection.right - selection.left + 1;
    const merge = sheet.merges.find((each) => each.top === selection.top && each.left === selection.left && each.bottom === selection.bottom && each.right === selection.right);
    const oneCell = (height === 1 && width === 1) || merge !== undefined;
    let down = 1;
    let across = 1;
    if (!oneCell && ((rows === 1 && cols === 1) || (height % rows === 0 && width % cols === 0))) {
      down = height / rows;
      across = width / cols;
    }
    const { top, left } = selection;
    plan = { dest: { top, left, bottom: top + down * rows - 1, right: left + across * cols - 1 }, top, left, rows, cols, down, across, transposed };
    if (merge && rows === 1 && cols === 1 && block.merges.length === 0) {
      plan.intoMerge = merge;
    }
  }
  if (plan.dest.bottom >= MAX_ROWS || plan.dest.right >= MAX_COLS || plan.top + plan.down * rows > MAX_ROWS || plan.left + plan.across * cols > MAX_COLS) {
    return PASTE_PAST_EDGE_MESSAGE;
  }
  if (plan.down * rows * plan.across * cols > MAX_COPY_CELLS) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  return plan;
}

/** Why a plan cannot be carried out on the sheet at `sheetIndex`, or null:
 * a merge crossing the destination's edge, part of an array, a locked
 * cell. */
export function pasteRefusal(model: SpreadsheetModel, sheetIndex: number, plan: PastePlan): string | null {
  const sheet = model.sheets[sheetIndex]!;
  if (sheet.merges.some((merge) => rangesIntersect(merge, plan.dest) && !contains(plan.dest, merge) && merge !== plan.intoMerge)) {
    return PASTE_MERGE_MESSAGE;
  }
  if (arraysCrossing(sheet, plan.dest)) {
    return PARTIAL_ARRAY_MESSAGE;
  }
  return firstLockedCell(model, sheetIndex, [plan.dest]) ? PROTECTED_CELL_MESSAGE : null;
}

/** Whether an array formula's range crosses the rectangle's edge. */
function arraysCrossing(sheet: Sheet, dest: GridRange): boolean {
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.arrayRange && rangesIntersect(cell.arrayRange, dest) && !contains(dest, cell.arrayRange)) {
        return true;
      }
    }
  }
  return false;
}

// --- formulas --------------------------------------------------------------------------------

/** Whether a formula names a sheet (or a workbook) not in `known`
 * (upper-cased names). */
export function namesOtherSheets(formula: string, known: ReadonlySet<string>): boolean {
  for (const token of tokenize(formula)) {
    if (token.kind === "external") {
      return true;
    }
    if (token.kind === "ref" && token.ref) {
      const { sheet, sheetEnd } = token.ref;
      if ((sheet !== undefined && !known.has(sheet.toUpperCase())) || (sheetEnd !== undefined && !known.has(sheetEnd.toUpperCase()))) {
        return true;
      }
    } else if (token.kind === "name" || token.kind === "function" || token.kind === "error") {
      const prefix = sheetPrefixOf(token.text);
      if (prefix !== "") {
        const { sheet, sheetEnd } = parseSheetPrefix(prefix);
        if (!known.has(sheet.toUpperCase()) || (sheetEnd !== undefined && !known.has(sheetEnd.toUpperCase()))) {
          return true;
        }
      }
    }
  }
  return false;
}

function shiftedPart(part: RefPart, dRow: number, dCol: number): RefPart | null {
  const row = part.row !== undefined && !part.rowAbs ? part.row + dRow : part.row;
  const col = part.col !== undefined && !part.colAbs ? part.col + dCol : part.col;
  if ((row !== undefined && (row < 0 || row >= MAX_ROWS)) || (col !== undefined && (col < 0 || col >= MAX_COLS))) {
    return null;
  }
  const moved: RefPart = { rowAbs: part.rowAbs, colAbs: part.colAbs };
  if (row !== undefined) {
    moved.row = row;
  }
  if (col !== undefined) {
    moved.col = col;
  }
  return moved;
}

/** One corner under Transpose: relative on both axes → its offset from the
 * host transposed; otherwise shifted by the host's displacement. */
function transposedPart(part: RefPart, from: CellRef, to: CellRef): RefPart | null {
  if (part.row !== undefined && part.col !== undefined && !part.rowAbs && !part.colAbs) {
    const row = to.row + (part.col - from.col);
    const col = to.col + (part.row - from.row);
    return row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS ? null : { row, col, rowAbs: false, colAbs: false };
  }
  return shiftedPart(part, to.row - from.row, to.col - from.col);
}

/** A formula hosted at `from` as pasted transposed at `to` (see the
 * module header); off-grid references become #REF!. */
export function transposeFormula(formula: string, from: CellRef, to: CellRef): string {
  return rewriteReferences(formula, (ref): ReferenceNode | null => {
    const start = transposedPart(ref.start, from, to);
    const end = ref.end ? transposedPart(ref.end, from, to) : undefined;
    if (start === null || end === null) {
      return null;
    }
    if (!end) {
      return { ...ref, start };
    }
    // Corners in order, as Excel shows a range.
    let a = start;
    let b = end;
    if (a.row !== undefined && b.row !== undefined && a.row > b.row) {
      [a, b] = [
        { ...a, row: b.row, rowAbs: b.rowAbs },
        { ...b, row: a.row, rowAbs: a.rowAbs },
      ];
    }
    if (a.col !== undefined && b.col !== undefined && a.col > b.col) {
      [a, b] = [
        { ...a, col: b.col, colAbs: b.colAbs },
        { ...b, col: a.col, colAbs: a.colAbs },
      ];
    }
    return { ...ref, start: a, end: b };
  });
}

// --- the paste -------------------------------------------------------------------------------

/** A source position with no value and no formula. */
function isBlank(cell: PasteCell | undefined): boolean {
  return !cell || (cell.v === null && cell.f === undefined && cell.r1c1 === undefined);
}

/** A style's effective format code, for date classification. */
function codeOf(style: CellStyle | undefined): string | undefined {
  if (!style) {
    return undefined;
  }
  return (style.builtinFormatId === undefined ? undefined : BUILTIN_NUMBER_FORMATS[style.builtinFormatId]) ?? style.numberFormat;
}

export interface PasteExtras {
  /** The source workbook's style id behind each block style (same-workbook
   * pastes; the view's in-memory copy). */
  styleIds?: readonly number[];
}

/** Pastes `block` onto the sheet `sheetId` at the selection `target` (see
 * the module header). */
export function pasteBlock(model: SpreadsheetModel, sheetId: string, target: PasteTarget, block: PasteBlock, options: PasteOptions, extras: PasteExtras = {}): PasteResult {
  const sheetIndex = model.sheets.findIndex((each) => each.id === sheetId);
  const sheet = model.sheets[sheetIndex];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  const plan = planPaste(sheet, target, block, options.transpose);
  if (typeof plan === "string") {
    return { ok: false, message: plan };
  }
  const refusal = pasteRefusal(model, sheetIndex, plan);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  const what = options.what;
  const replacesFormats = what === "all" || what === "formats";
  // Merges kept by the paste: positions they cover, but their anchors, are
  // not written.
  const kept = replacesFormats ? (plan.intoMerge ? [plan.intoMerge] : []) : sheet.merges.filter((merge) => contains(plan.dest, merge));
  const covered = (row: number, col: number): boolean =>
    kept.some((merge) => row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right && (row !== merge.top || col !== merge.left));
  const byPosition = new Map<number, PasteCell>();
  for (const cell of block.cells) {
    byPosition.set(cell.r * block.cols + cell.c, cell);
  }
  // The source cell behind each destination offset in a tile.
  const sourceAt = (i: number, j: number): { cell: PasteCell | undefined; r: number; c: number } => {
    const r = plan.transposed ? j : i;
    const c = plan.transposed ? i : j;
    return { cell: byPosition.get(r * block.cols + c), r, c };
  };
  let formulasAsValues = 0;

  if (block.source === "text") {
    const entries: { row: number; col: number; entry: CellEntry | null }[] = [];
    forEachDestination(plan, (row, col, i, j) => {
      const { cell } = sourceAt(i, j);
      if ((options.skipBlanks && isBlank(cell)) || covered(row, col)) {
        return;
      }
      // Values (and Values and number formats) drop formulas: a formula
      // field pastes as its text.
      let entry = (what === "values" || what === "values-number-formats" ? cell?.asText : undefined) ?? cell?.entry ?? null;
      if (entry && what === "values") {
        const { numberFormat: _format, builtinFormatId: _builtin, ...plain } = entry;
        entry = plain;
      }
      entries.push({ row, col, entry });
    });
    writeEntries(model, sheetIndex, entries);
    if (replacesFormats && !plan.intoMerge) {
      // Paste (All) replaces the destination's merges, as for a payload:
      // plain text brings none, so every merge inside is gone.
      const current = model.sheets[sheetIndex]!;
      const merges = current.merges.filter((merge) => !contains(plan.dest, merge));
      if (merges.length !== current.merges.length) {
        model.sheets = model.sheets.map((each, index) => (index === sheetIndex ? { ...current, merges, dirty: true as const } : each));
      }
    }
    markDirty(model, sheetId, sheet);
    return { ok: true, range: writtenRange(plan), formulasAsValues };
  }

  const sameWorkbook = block.workbookId !== "" && block.workbookId === model.id;
  const knownSheets = new Set(model.sheets.map((each) => each.name.toUpperCase()));
  const dateShift = block.dateSystem === model.dateSystem ? 0 : model.dateSystem === "1904" ? -DATE_SYSTEM_OFFSET : DATE_SYSTEM_OFFSET;
  const mapped: (number | undefined)[] = [];
  const styleId = (index: number): number => {
    let id = mapped[index];
    if (id === undefined) {
      const style = block.styles[index]!;
      const own = sameWorkbook ? extras.styleIds?.[index] : undefined;
      id = own !== undefined && extras.styleIds?.length === block.styles.length && model.styles[own] !== undefined && styleKey(model.styles[own]!) === styleKey(style) ? own : internStyle(model, style);
      mapped[index] = id;
    }
    return id;
  };
  const blankStyle: number | "inherit" = plan.whole ? "inherit" : 0;
  const edit = new SheetEdit(sheet);
  const numberFormatIds = new Map<string, number>();

  /** The change one source position makes at (row, col). */
  const changeFor = (row: number, col: number, cell: PasteCell | undefined, r: number, c: number): CellChange | null => {
    if (options.skipBlanks && isBlank(cell)) {
      return null;
    }
    const change: CellChange = {};
    if (what !== "formats") {
      let value = cell?.v ?? null;
      const sourceStyle = cell?.s === undefined ? undefined : block.styles[cell.s];
      if (dateShift !== 0 && typeof value === "number" && dateKindOf(codeOf(sourceStyle)) === "date") {
        value += dateShift;
      }
      let formula: string | undefined;
      let arrayRange: GridRange | undefined;
      if ((what === "all" || what === "formulas") && cell && (cell.f !== undefined || cell.r1c1 !== undefined)) {
        const from = { row: block.origin.row + r, col: block.origin.col + c };
        const to = { row, col };
        let converted: string | null = null;
        let otherSheets = false;
        if (cell.r1c1 !== undefined) {
          // Google Sheets: R1C1 converted at its destination — another
          // workbook's formula, under the same missing-sheet rule.
          const a1 = r1c1ToA1(cell.r1c1, to, plan.transposed);
          otherSheets = a1 !== null && namesOtherSheets(a1, knownSheets);
          converted = a1 === null || otherSheets || a1.length > MAX_CLIPBOARD_FORMULA_LENGTH ? null : a1;
        } else if (!sameWorkbook && namesOtherSheets(cell.f!, knownSheets)) {
          otherSheets = true;
        } else {
          converted = plan.transposed ? transposeFormula(cell.f!, from, to) : shiftFormula(cell.f!, row - from.row, col - from.col);
        }
        if (converted === null) {
          formulasAsValues += otherSheets ? 1 : 0;
        } else {
          formula = converted;
          if (cell.a && !plan.transposed && what === "all") {
            arrayRange = { top: row, left: col, bottom: row + cell.a.bottom - cell.a.top, right: col + cell.a.right - cell.a.left };
          }
        }
      }
      change.content = { value };
      if (formula !== undefined) {
        change.content.formula = formula;
      }
      if (arrayRange) {
        change.content.arrayRange = arrayRange;
      }
      if (what === "all" && cell?.t && formula === undefined && typeof value === "string") {
        change.content.runs = structuredClone(cell.t);
      }
    }
    if (replacesFormats) {
      change.style = cell?.s !== undefined ? styleId(cell.s) : blankStyle;
    } else if (what === "values-number-formats") {
      const source = cell?.s === undefined ? {} : block.styles[cell.s]!;
      const base = edit.cell(row, col)?.styleId ?? inheritedStyleId(sheet, row, col);
      const key = `${base}|${source.numberFormat ?? ""}|${source.builtinFormatId ?? ""}`;
      let id = numberFormatIds.get(key);
      if (id === undefined) {
        id = internStyle(model, patchStyle(cellStyle(model, base), { numberFormat: source.numberFormat ?? null, builtinFormatId: source.builtinFormatId ?? null }));
        numberFormatIds.set(key, id);
      }
      change.style = id;
    }
    return change;
  };

  forEachDestination(plan, (row, col, i, j) => {
    if (covered(row, col)) {
      return;
    }
    const { cell, r, c } = sourceAt(i, j);
    const change = changeFor(row, col, cell, r, c);
    if (change) {
      edit.change(row, col, change);
    }
  });
  if (plan.whole) {
    clearOutside(sheet, plan, edit, (row, col, existing) => {
      const change = changeFor(row, col, undefined, 0, 0);
      return change ? cellWith(existing, col, change, inheritedStyleId(sheet, row, col)) : existing;
    });
  }
  if (plan.whole && replacesFormats) {
    wholeRowProps(plan, block, edit, styleId);
  }
  const rows = edit.rows();
  let merges = sheet.merges;
  if (replacesFormats && !plan.intoMerge) {
    merges = [...sheet.merges.filter((merge) => !contains(plan.dest, merge)), ...tiledMerges(plan, block)];
  }
  const columns = plan.whole && replacesFormats && (plan.whole === "cols" || plan.whole === "sheet") ? wholeColumns(sheet, plan, block, styleId) : sheet.columns;
  if (rows !== sheet.rows || merges !== sheet.merges || columns !== sheet.columns || model.sheets[sheetIndex] !== sheet) {
    const current = model.sheets[sheetIndex]!;
    model.sheets = model.sheets.map((each, index) => (index === sheetIndex ? { ...current, rows, merges, columns, dirty: true as const } : each));
  }
  return { ok: true, range: writtenRange(plan), formulasAsValues };
}

/** Every destination position of a plan, with its offset in its tile. */
function forEachDestination(plan: PastePlan, visit: (row: number, col: number, i: number, j: number) => void): void {
  for (let tileRow = 0; tileRow < plan.down; tileRow += 1) {
    for (let i = 0; i < plan.rows; i += 1) {
      const row = plan.top + tileRow * plan.rows + i;
      for (let tileCol = 0; tileCol < plan.across; tileCol += 1) {
        for (let j = 0; j < plan.cols; j += 1) {
          visit(row, plan.left + tileCol * plan.cols + j, i, j);
        }
      }
    }
  }
}

/** For whole rows/columns: every stored cell of the destination band
 * outside the written block, through `next` (a blank source position). */
function clearOutside(sheet: Sheet, plan: PastePlan, edit: SheetEdit, next: (row: number, col: number, existing: ReturnType<SheetEdit["cell"]>) => ReturnType<SheetEdit["cell"]>): void {
  const written = writtenRange(plan);
  for (const row of sheet.rows) {
    if (row.index < plan.dest.top || row.index > plan.dest.bottom) {
      continue;
    }
    for (const cell of row.cells) {
      if (cell.col < plan.dest.left || cell.col > plan.dest.right) {
        continue;
      }
      if (row.index >= written.top && row.index <= written.bottom && cell.col >= written.left && cell.col <= written.right) {
        continue;
      }
      const replaced = next(row.index, cell.col, cell);
      if (replaced !== cell) {
        edit.set(row.index, cell.col, replaced);
      }
    }
  }
}

/** Whole rows: each destination row's properties replaced by its source's. */
function wholeRowProps(plan: PastePlan, block: PasteBlock, edit: SheetEdit, styleId: (index: number) => number): void {
  if (plan.whole !== "rows" && plan.whole !== "sheet") {
    return;
  }
  const props = new Map((block.rowProps ?? []).map((entry) => [entry.r, entry] as const));
  for (let r = 0; r < block.rows; r += 1) {
    const source = props.get(r);
    const row: { heightPt?: number; customHeight?: true; hidden?: true; styleId?: number } = {};
    if (source?.heightPt !== undefined) {
      row.heightPt = source.heightPt;
    }
    if (source?.customHeight) {
      row.customHeight = true;
    }
    if (source?.hidden) {
      row.hidden = true;
    }
    if (source?.s !== undefined) {
      row.styleId = styleId(source.s);
    }
    edit.setRowProps(plan.top + r, row);
  }
}

/** Whole columns: the destination columns' spans replaced by the source's. */
function wholeColumns(sheet: Sheet, plan: PastePlan, block: PasteBlock, styleId: (index: number) => number): ColumnSpan[] {
  const first = plan.left;
  const last = plan.left + block.cols - 1;
  const out: ColumnSpan[] = [];
  for (const span of sheet.columns) {
    if (span.max < first || span.min > last) {
      out.push(span);
      continue;
    }
    if (span.min < first) {
      out.push({ ...span, max: first - 1 });
    }
    if (span.max > last) {
      out.push({ ...span, min: last + 1 });
    }
  }
  for (const props of block.colProps ?? []) {
    const span: ColumnSpan = { min: first + props.c, max: first + props.c };
    if (props.widthPx !== undefined) {
      span.widthPx = props.widthPx;
    }
    if (props.hidden) {
      span.hidden = true;
    }
    if (props.s !== undefined) {
      span.styleId = styleId(props.s);
    }
    out.push(span);
  }
  return joinSpans(out.toSorted((a, b) => a.min - b.min));
}

/** The block's merges at every tile (transposed with the block). */
function tiledMerges(plan: PastePlan, block: PasteBlock): GridRange[] {
  const out: GridRange[] = [];
  for (let tileRow = 0; tileRow < plan.down; tileRow += 1) {
    for (let tileCol = 0; tileCol < plan.across; tileCol += 1) {
      const top = plan.top + tileRow * plan.rows;
      const left = plan.left + tileCol * plan.cols;
      for (const merge of block.merges) {
        out.push(
          plan.transposed
            ? { top: top + merge.left, left: left + merge.top, bottom: top + merge.right, right: left + merge.bottom }
            : { top: top + merge.top, left: left + merge.left, bottom: top + merge.bottom, right: left + merge.right },
        );
      }
    }
  }
  return out;
}

/** Marks a sheet dirty when its rows changed (the entry path). */
function markDirty(model: SpreadsheetModel, sheetId: string, before: Sheet): void {
  model.sheets = model.sheets.map((each) => (each.id === sheetId && each.rows !== before.rows && !each.dirty ? { ...each, dirty: true as const } : each));
}
