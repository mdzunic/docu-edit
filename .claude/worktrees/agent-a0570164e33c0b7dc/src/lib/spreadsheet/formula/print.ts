// Canonical text out of the formula model (SPEC-119 "Canonical text",
// "Sheet-name quoting", "The formula bar"):
//
// - printFormula(node): the canonical text of an AST — function names as the
//   node holds them, references through formatReference, no whitespace but
//   the intersection space, and the minimal parentheses: a parsed AST keeps
//   its own `paren` nodes, so parse(print(parse(t))) equals parse(t) (spans
//   aside); a built AST gets parentheses exactly where precedence, left
//   associativity or an argument's comma needs them. It is a test oracle and
//   a builder (SPEC-125, SPEC-140) — no rewrite in this queue is
//   print(parse()), which would discard the user's whitespace.
// - Sheet-name quoting: a name is written bare iff it matches
//   ^[\p{L}_][\p{L}\p{N}_.]*$ and is not a cell reference (A1, XFD1), not
//   R1C1-shaped (R, C, R1, C1, RC, R1C1, any case) and not TRUE/FALSE;
//   otherwise it is wrapped in ' with every ' doubled. A 3-D prefix is quoted
//   whole (`'Q1 2026:Q4 2026'!`). The ODS translator and every rename use the
//   same predicate.
// - formulaBarText(sheet, row, col, style, ctx): what the read-only formula
//   bar shows — "=" + formula for a formula cell; "{=" + anchor formula + "}"
//   for every cell of a legacy array range, members included; a dynamic
//   array's anchor "=" + formula, its members the anchor's formula marked
//   `spillMember` (the bar renders it muted); a data table range
//   `{=TABLE(r1,r2)}` (one-dimensional: `{=TABLE(r1,)}` with dtr="1", else
//   `{=TABLE(,r1)}`); an unparsed formula as stored ("=" + the xlsx `<f>`
//   text; an ods `table:formula` from its first "="); and SPEC-118's editText
//   for every other cell.
import { colToLetters, getCell, parseA1 } from "../address";
import type { Cell, CellStyle, Sheet, SpreadsheetFormat } from "../model";
import { editText, type FormatContext } from "../numberFormat/render";
import { formulaAnchorAt, startTagAttribute } from "./anchors";
import type { ArrayElement, BinaryOp, FormulaNode } from "./ast";
import type { RefPart, ReferenceNode } from "./references";

// --- sheet names --------------------------------------------------------------------

const BARE_SHEET_NAME = /^[\p{L}_][\p{L}\p{N}_.]*$/u;
const R1C1_SHAPED = /^(?:R\d*C\d*|R\d*|C\d*)$/i;

/** Whether a sheet name must be quoted in a reference (see the header). */
export function sheetNameNeedsQuotes(name: string): boolean {
  if (!BARE_SHEET_NAME.test(name) || parseA1(name) !== null || R1C1_SHAPED.test(name)) {
    return true;
  }
  const upper = name.toUpperCase();
  return upper === "TRUE" || upper === "FALSE";
}

function escapeQuotes(name: string): string {
  return name.replaceAll("'", "''");
}

/** A sheet name as a reference writes it: bare, or quoted with ' doubled. */
export function quoteSheetName(name: string): string {
  return sheetNameNeedsQuotes(name) ? `'${escapeQuotes(name)}'` : name;
}

/** A sheet prefix, `!` included: `Sheet1!`, `'My Sheet'!`, `Jan:Mar!`,
 * `'Q1 2026:Q4 2026'!` (a 3-D prefix is quoted whole when either end needs
 * it). */
export function formatSheetPrefix(sheet: string, sheetEnd?: string): string {
  if (sheetEnd === undefined) {
    return `${quoteSheetName(sheet)}!`;
  }
  if (sheetNameNeedsQuotes(sheet) || sheetNameNeedsQuotes(sheetEnd)) {
    return `'${escapeQuotes(sheet)}:${escapeQuotes(sheetEnd)}'!`;
  }
  return `${sheet}:${sheetEnd}!`;
}

function partText(part: RefPart): string {
  let text = "";
  if (part.col !== undefined) {
    text += `${part.colAbs ? "$" : ""}${colToLetters(part.col)}`;
  }
  if (part.row !== undefined) {
    text += `${part.rowAbs ? "$" : ""}${part.row + 1}`;
  }
  return text;
}

/** A reference's canonical text: `$` markers kept, letters upper-case, the
 * sheet prefix quoted per the rule; a whole column or row is always a range
 * (`A:A`, `1:1`). */
export function formatReference(ref: ReferenceNode): string {
  const prefix = ref.sheet === undefined ? "" : formatSheetPrefix(ref.sheet, ref.sheetEnd);
  const whole = ref.start.row === undefined || ref.start.col === undefined;
  const end = ref.end ?? (whole ? ref.start : undefined);
  return prefix + partText(ref.start) + (end ? `:${partText(end)}` : "");
}

// --- printFormula -------------------------------------------------------------------

/** Binding power — higher binds tighter (SPEC-119 "Precedence"). */
const BINARY_POWER: Readonly<Record<BinaryOp, number>> = {
  ":": 10,
  " ": 9,
  ",": 8,
  "^": 5,
  "*": 4,
  "/": 4,
  "+": 3,
  "-": 3,
  "&": 2,
  "=": 1,
  "<>": 1,
  "<": 1,
  "<=": 1,
  ">": 1,
  ">=": 1,
};
const PREFIX_POWER = 7;
const PERCENT_POWER = 6;
const ATOM_POWER = 11;

function powerOf(node: FormulaNode): number {
  if (node.kind === "binary") {
    return BINARY_POWER[node.op];
  }
  if (node.kind === "unary") {
    return node.op === "%" ? PERCENT_POWER : PREFIX_POWER;
  }
  return ATOM_POWER;
}

function numberText(value: number): string {
  return String(value).toUpperCase();
}

function stringText(value: string): string {
  return `"${value.replaceAll('"', '""')}"`;
}

function elementText(element: ArrayElement): string {
  if (typeof element === "number") {
    return numberText(element);
  }
  if (typeof element === "string") {
    return stringText(element);
  }
  if (typeof element === "boolean") {
    return element ? "TRUE" : "FALSE";
  }
  return element.error;
}

/** `node` printed; `inArgument` says a bare comma would split an argument
 * list here, so a union must be parenthesised. */
function printNode(node: FormulaNode, inArgument: boolean): string {
  const wrap = (inner: FormulaNode, needed: boolean): string =>
    needed ? `(${printNode(inner, false)})` : printNode(inner, inArgument);
  switch (node.kind) {
    case "number":
      return numberText(node.value);
    case "string":
      return stringText(node.value);
    case "boolean":
      return node.value ? "TRUE" : "FALSE";
    case "error":
      return node.text;
    case "array":
      return `{${node.rows.map((row) => row.map(elementText).join(",")).join(";")}}`;
    case "ref":
      return formatReference(node.ref);
    case "name":
      return node.sheet === undefined ? node.name : formatSheetPrefix(node.sheet) + node.name;
    case "function":
      return `${node.name}(${node.args.map((arg) => printNode(arg, true)).join(",")})`;
    case "missing":
      return "";
    case "unary":
      if (node.op === "%") {
        return `${wrap(node.operand, powerOf(node.operand) < PERCENT_POWER)}%`;
      }
      return node.op + wrap(node.operand, powerOf(node.operand) < PREFIX_POWER);
    case "binary": {
      if (inArgument && node.op === ",") {
        return `(${printNode(node, false)})`;
      }
      const power = BINARY_POWER[node.op];
      return wrap(node.left, powerOf(node.left) < power) + node.op + wrap(node.right, powerOf(node.right) <= power);
    }
    case "paren":
      return `(${printNode(node.inner, false)})`;
    case "structured":
    case "external":
      return node.text;
  }
}

/** The canonical text of an AST (see the module header). */
export function printFormula(node: FormulaNode): string {
  return printNode(node, false);
}

// --- the formula bar ----------------------------------------------------------------

/** What the formula bar's context adds to SPEC-118's: the workbook's format,
 * which decides how an unparsed formula is shown (default xlsx). */
export interface FormulaBarContext extends FormatContext {
  format?: SpreadsheetFormat;
}

/** The bar's text, and whether it is a dynamic-array member's (muted). */
export interface FormulaBarContent {
  text: string;
  spillMember: boolean;
}

function isOn(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

/** A data table's text from its retained `<f t="dataTable">` element. */
export function dataTableText(xml: string): string {
  const input = (index: 1 | 2): string =>
    isOn(startTagAttribute(xml, `del${index}`)) ? "#REF!" : (startTagAttribute(xml, `r${index}`) ?? "");
  if (isOn(startTagAttribute(xml, "dt2D"))) {
    return `{=TABLE(${input(1)},${input(2)})}`;
  }
  return isOn(startTagAttribute(xml, "dtr")) ? `{=TABLE(${input(1)},)}` : `{=TABLE(,${input(1)})}`;
}

/** A formula as the bar writes it: "=" + canonical text, or an unparsed
 * formula as stored. */
function storedFormula(cell: Cell & { formula: string }, format: SpreadsheetFormat): string {
  if (cell.formulaUnparsed && format === "ods") {
    const equals = cell.formula.indexOf("=");
    return equals < 0 ? cell.formula : cell.formula.slice(equals);
  }
  return `=${cell.formula}`;
}

/** The formula a position shows — its own, its array's or its data table's
 * — or undefined for a position no formula covers. */
export function formulaTextAt(sheet: Sheet, row: number, col: number, format: SpreadsheetFormat = "xlsx"): FormulaBarContent | undefined {
  const cell = getCell(sheet, row, col);
  if (cell?.formula !== undefined) {
    const text = storedFormula(cell as Cell & { formula: string }, format);
    return { text: cell.arrayRange && !cell.dynamicArray ? `{${text}}` : text, spillMember: false };
  }
  if (cell?.dataTableXml !== undefined) {
    return { text: dataTableText(cell.dataTableXml), spillMember: false };
  }
  const anchor = formulaAnchorAt(sheet, row, col);
  if (anchor?.kind === "dataTable") {
    return { text: dataTableText(anchor.cell.dataTableXml!), spillMember: false };
  }
  if (anchor?.cell.formula !== undefined) {
    const text = storedFormula(anchor.cell as Cell & { formula: string }, format);
    return anchor.cell.dynamicArray ? { text, spillMember: true } : { text: `{${text}}`, spillMember: false };
  }
  return undefined;
}

/** The formula bar's text and muting (see the module header). `style` is
 * the position's effective style, passed to editText. */
export function formulaBarContent(sheet: Sheet, row: number, col: number, style: CellStyle | undefined, ctx: FormulaBarContext): FormulaBarContent {
  const formula = formulaTextAt(sheet, row, col, ctx.format);
  if (formula) {
    return formula;
  }
  const cell = getCell(sheet, row, col);
  return { text: cell ? editText(cell, { ...ctx, style }) : "", spillMember: false };
}

/** The formula bar's text (see the module header). */
export function formulaBarText(sheet: Sheet, row: number, col: number, style: CellStyle | undefined, ctx: FormulaBarContext): string {
  return formulaBarContent(sheet, row, col, style, ctx).text;
}
