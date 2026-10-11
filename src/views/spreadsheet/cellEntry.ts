// What a cell edit writes (SPEC-122 "Input parsing", "Cell commands",
// "Protection", "Row autofit"): the editor's text for a cell, the CellEntry a
// committed text becomes, the refusals checked before the editor opens or a
// Delete or Fill runs, and the three writes — each ONE undo step through
// SPEC-121's record/commit with the touched rows refitted inside it.
//
// - Edit text: "=" + the formula for a formula cell (an unparsed one as
//   stored, an array anchor without its braces), else SPEC-118's editText
//   through the effective style (a quote-prefixed string keeps its "'").
// - Commit: the text passes sanitizeXmlText, then parseInput in the user's
//   locale and the workbook's date system with the cell's effective format.
//   A formula over MAX_FORMULA_LENGTH characters, or one SPEC-119's
//   parseFormula rejects once its unbalanced trailing "(" are closed, is an
//   error with the caret offset to show; a valid one is stored canonical
//   (SPEC-119's printFormula). A text holding a line break asks for wrap
//   (lib's entryFromInput; setCellEntry interns it).
// - Each write refits the rows it touched (lib's refitRows given the rows
//   before it) inside its own undo step.
// - Refusals: a locked cell on a protected sheet ("This cell is on a
//   protected sheet") before part of an array ("You can't change part of an
//   array" — lib's arrayRangeTouched lets the anchor alone be edited).
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { formatA1, getCell } from "../../lib/spreadsheet/address";
import { refitRows, type CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import {
  ARRAY_PART_MESSAGE,
  arrayRangeTouched,
  clearContents,
  entryFromInput,
  fillWithEntry,
  firstLockedCell,
  PROTECTED_CELL_MESSAGE,
  setCellEntry,
  type CellEntry,
} from "../../lib/spreadsheet/cells";
import { sanitizeXmlText } from "../../lib/core/xmlText";
import { FormulaSyntaxError, parseFormula } from "../../lib/spreadsheet/formula/parse";
import { formulaTextAt, printFormula } from "../../lib/spreadsheet/formula/print";
import { tokenize } from "../../lib/spreadsheet/formula/tokens";
import type { CellStyle, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { formatCodeOf } from "../../lib/spreadsheet/numberFormat/builtin";
import { parseInput } from "../../lib/spreadsheet/numberFormat/input";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { editText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export const FORMULA_ERROR_MESSAGE = "There's a problem with this formula";
export const FORMULA_TOO_LONG_MESSAGE = "Formulas can be at most 8,192 characters";
export { ARRAY_PART_MESSAGE, PROTECTED_CELL_MESSAGE };
/** Excel's limit on a formula's length, "=" excluded. */
export const MAX_FORMULA_LENGTH = 8192;
export const FILL_LABEL = "Fill";
export const CLEAR_CONTENTS_LABEL = "Clear contents";

/** The undo label of a commit: `Typing "<first 20 characters>" in B3`. */
export function typingLabel(text: string, at: CellRef): string {
  return `Typing "${Array.from(text).slice(0, 20).join("")}" in ${formatA1(at)}`;
}

/** The editor's text for a position (see the module header). */
export function cellEditText(model: SpreadsheetModel, sheet: Sheet, at: CellRef, style: CellStyle | undefined, ctx: FormatContext): string {
  const cell = getCell(sheet, at.row, at.col);
  if (!cell) {
    return "";
  }
  if (cell.formula !== undefined) {
    const shown = formulaTextAt(sheet, at.row, at.col, model.format)?.text ?? `=${cell.formula}`;
    const braced = cell.arrayRange !== undefined && !cell.dynamicArray && shown.startsWith("{") && shown.endsWith("}");
    return braced ? shown.slice(1, -1) : shown;
  }
  return editText(cell, { ...ctx, style });
}

/** A formula body with its unbalanced "(" closed, as Excel offers. */
export function closeOpenParens(body: string): string {
  let depth = 0;
  for (const token of tokenize(body)) {
    if (token.kind === "open") {
      depth += 1;
    } else if (token.kind === "close") {
      depth = Math.max(0, depth - 1);
    }
  }
  return depth > 0 ? body + ")".repeat(depth) : body;
}

export type EntryOutcome = { entry: CellEntry } | { error: string; caret: number };

export interface EntryContext {
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
  /** The cell's effective style. */
  style: CellStyle;
}

/** The entry a committed text writes, or why it cannot (see the header). */
export function entryFromText(text: string, ctx: EntryContext): EntryOutcome {
  const clean = sanitizeXmlText(text);
  const parsed = parseInput(clean, ctx.locale, ctx.dateSystem, { existingFormat: formatCodeOf(ctx.style, ctx.locale) });
  if (parsed.formula === undefined) {
    return { entry: entryFromInput(parsed) };
  }
  const body = clean.startsWith("=") ? clean.slice(1) : parsed.formula;
  if (body.length > MAX_FORMULA_LENGTH) {
    return { error: FORMULA_TOO_LONG_MESSAGE, caret: clean.length };
  }
  try {
    return { entry: entryFromInput({ ...parsed, value: null, formula: printFormula(parseFormula(closeOpenParens(body))) }) };
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return { error: FORMULA_ERROR_MESSAGE, caret: Math.min(clean.length, 1 + error.offset) };
    }
    throw error;
  }
}

/** A formula body as a commit stores it (canonical, unbalanced "(" closed);
 * null when it is over MAX_FORMULA_LENGTH or no formula — SPEC-125's
 * plain-text paste reads its fields through this. */
export function canonicalFormula(body: string): string | null {
  if (body.length > MAX_FORMULA_LENGTH) {
    return null;
  }
  try {
    return printFormula(parseFormula(closeOpenParens(body)));
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return null;
    }
    throw error;
  }
}

/** Why `ranges` cannot be changed, or null (see the header). */
export function refusalFor(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): string | null {
  if (firstLockedCell(model, sheetIndex, ranges)) {
    return PROTECTED_CELL_MESSAGE;
  }
  return arrayRangeTouched(model, sheetIndex, ranges) ? ARRAY_PART_MESSAGE : null;
}

/** Whether protection lets a grid sheet's cell be edited (SPEC-127: the fx
 * button and Shift+F3 do nothing on a cell the editor refuses this way). */
export function canEditCell(model: SpreadsheetModel | null, sheetIndex: number, sheet: Sheet | null, at: CellRef): boolean {
  return model !== null && sheet?.kind === "grid" && !firstLockedCell(model, sheetIndex, [{ top: at.row, left: at.col, bottom: at.row, right: at.col }]);
}

/** What every write needs: the live model, the sheet, the undo pipeline,
 * the measurer and format context autofit takes. */
export interface CellWriteContext {
  model: SpreadsheetModel;
  sheetIndex: number;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  measurer: CellTextMeasurer;
  format: FormatContext;
}

function oneStep(ctx: CellWriteContext, label: string, ranges: GridRange[], mutate: () => void): void {
  const sheetId = ctx.model.sheets[ctx.sheetIndex]?.id;
  if (sheetId === undefined) {
    return;
  }
  ctx.undo.record(label, { sheetId });
  const before = ctx.model.sheets[ctx.sheetIndex]!.rows;
  mutate();
  // Only the rows the command touched refit: an unchanged entry refits (and
  // records) nothing, and an untouched row keeps the height its file gave it.
  refitRows(ctx.model, sheetId, ranges, ctx.measurer, ctx.format, before);
  ctx.undo.commit();
}

/** A commit: one step `Typing "…" in B3` (an unchanged entry records nothing —
 * commit() pops a step that changed nothing). */
export function writeEntry(ctx: CellWriteContext, at: CellRef, entry: CellEntry, text: string): void {
  const range = { top: at.row, left: at.col, bottom: at.row, right: at.col };
  oneStep(ctx, typingLabel(text, at), [range], () => setCellEntry(ctx.model, ctx.sheetIndex, at, entry));
}

/** Ctrl/Cmd+Enter: one step "Fill" over the selection. */
export function writeFill(ctx: CellWriteContext, ranges: GridRange[], origin: CellRef, entry: CellEntry): void {
  oneStep(ctx, FILL_LABEL, ranges, () => fillWithEntry(ctx.model, ctx.sheetIndex, ranges, origin, entry));
}

/** Delete/Backspace: one step "Clear contents" over the selection. */
export function writeClear(ctx: CellWriteContext, ranges: GridRange[]): void {
  oneStep(ctx, CLEAR_CONTENTS_LABEL, ranges, () => clearContents(ctx.model, ctx.sheetIndex, ranges));
}
