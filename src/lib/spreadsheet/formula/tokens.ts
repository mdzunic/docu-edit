// The formula tokenizer (SPEC-119 "Tokens", "Token rules") — THE normative
// source for how canonical formula text splits into tokens. Every later
// rewrite (shiftFormula, renameSheetInFormula, SPEC-124's adjustFormula,
// SPEC-133's table rewrites) works on this token stream, never on a printed
// AST, because the stream keeps the user's whitespace:
//
// - Tokenizing is TOTAL: it never throws. A character no rule claims becomes
//   an `unknown` token, an unterminated string or structured reference runs
//   to the end of the text (`unterminated: true`), and the parser — not the
//   tokenizer — decides what is a syntax error.
// - Every token carries the whitespace (space, tab, CR, LF) immediately
//   before it as `lead`; the closing `end` token holds the trailing
//   whitespace. So joinTokens(tokenize(t)) === t for every string t.
// - A reference — cell, range, whole column (`A:A`, `$A:$C`), whole row
//   (`1:1`), sheet-qualified (`Sheet1!A1`, `'My Sheet'!A1:B2`,
//   `'O''Brien'!A1`) or 3-D (`Sheet1:Sheet3!A1`, `'Q1 2026:Q4 2026'!C3`) —
//   is ONE `ref` token carrying its parsed ReferenceNode. The parts join
//   greedily with no whitespace around `:`; `Sheet1!A1:Sheet1!B2` is two
//   refs and a `:` operator, and a lone `A` is a name.
// - An identifier (a letter, `_` or `\`, then letters, digits, `_` and `.`)
//   immediately followed by `(` is a `function` whatever it spells (`LOG10(`);
//   otherwise it is a `ref` when it is a cell inside A1..XFD1048576, a
//   `boolean` for TRUE/FALSE in any case, and a `name` otherwise (`XFE1`,
//   `A0`, `R1C1`, `Rate`, `Sheet1!Rate`).
// - Errors: the fifteen codes, case-insensitively; `Sheet1!#REF!` and
//   `#REF!A1` (a reference into a deleted sheet) are each one `error` token.
// - `[1]Sheet1!A1`, `[1]!Name` and `'[1]My Sheet'!A1` are one opaque
//   `external` token, as is an OpenFormula external reference kept verbatim
//   by the ODS translator (`['file:///b.ods'#$Sheet1.A1]`).
// - `Table1[[#This Row],[Amount]]` and `[@Qty]` are one `structured` token,
//   consumed by balanced brackets; `'` escapes the next character.
// - A whitespace run containing a space, between a reference-ending token
//   (ref, name, structured, external, `)`) and a reference-starting one (ref,
//   name, function, structured, external, `(`), is the intersection
//   `operator` — its text is the run's LAST space; the run before it is its
//   lead, the run after it the next token's lead.
// - A comma is `argSep` when the innermost open bracket is a function call's
//   `(`, `arrayColSep` inside `{…}`, and `union` otherwise (top level or a
//   grouping paren); `;` is `arrayRowSep` inside `{…}` and `unknown`
//   elsewhere. A sign directly before a number inside `{…}` belongs to the
//   number (`{1,-2}`).
import { lettersToCol } from "../address";
import { ERROR_CODES, MAX_ROWS } from "../model";
import type { RefPart, ReferenceNode } from "./references";

export type TokenKind =
  | "number"
  | "string"
  | "boolean"
  | "error"
  | "ref"
  | "name"
  | "function"
  | "structured"
  | "external"
  | "operator" // + - * / ^ & = <> < <= > >= % : and " " (intersection)
  | "open"
  | "close"
  | "argSep"
  | "union"
  | "arrayOpen"
  | "arrayClose"
  | "arrayColSep"
  | "arrayRowSep"
  | "unknown"
  | "end"; // "end" holds trailing whitespace as its lead

export interface Token {
  kind: TokenKind;
  /** The token's own source text. */
  text: string;
  /** Whitespace (space, tab, CR, LF) immediately before it. */
  lead: string;
  /** UTF-16 offset of `text` in the source. */
  start: number;
  /** The parsed reference of a `ref` token. */
  ref?: ReferenceNode;
  /** A string or structured token that ran off the end of the text. */
  unterminated?: true;
}

/** Kinds after which a space may be the intersection operator. */
const REFERENCE_ENDING: ReadonlySet<TokenKind> = new Set(["ref", "name", "structured", "external", "close"]);
/** Kinds before which a space may be the intersection operator. */
const REFERENCE_STARTING: ReadonlySet<TokenKind> = new Set(["ref", "name", "function", "structured", "external", "open"]);

const UPPER_ERROR_CODES = ERROR_CODES.map((code) => code.toUpperCase());

const LETTER = /\p{L}/u;
const LETTER_OR_NUMBER = /[\p{L}\p{N}]/u;

function isWhitespace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

function isAsciiLetter(code: number): boolean {
  return (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a);
}

function isDigit(code: number): boolean {
  return code >= 0x30 && code <= 0x39;
}

/** UTF-16 units of the code point at `i` (2 for a surrogate pair). */
function unitsAt(text: string, i: number): number {
  const code = text.charCodeAt(i);
  return code >= 0xd800 && code <= 0xdbff && i + 1 < text.length ? 2 : 1;
}

/** A letter, `_` or `\` at `i`. */
export function isIdentifierStart(text: string, i: number): boolean {
  if (i >= text.length) {
    return false;
  }
  const code = text.charCodeAt(i);
  if (code < 0x80) {
    return isAsciiLetter(code) || code === 0x5f || code === 0x5c;
  }
  return LETTER.test(String.fromCodePoint(text.codePointAt(i)!));
}

/** A letter, digit, `_` or `.` at `i`. */
export function isIdentifierPart(text: string, i: number): boolean {
  if (i >= text.length) {
    return false;
  }
  const code = text.charCodeAt(i);
  if (code < 0x80) {
    return isAsciiLetter(code) || isDigit(code) || code === 0x5f || code === 0x2e;
  }
  return LETTER_OR_NUMBER.test(String.fromCodePoint(text.codePointAt(i)!));
}

/** The end of the identifier starting at `i` (which must start one). */
export function identifierEnd(text: string, i: number): number {
  let j = i + unitsAt(text, i);
  while (isIdentifierPart(text, j)) {
    j += unitsAt(text, j);
  }
  return j;
}

// --- areas ----------------------------------------------------------------------

interface ColumnScan {
  end: number;
  col: number;
  abs: boolean;
}

interface RowScan {
  end: number;
  row: number;
  abs: boolean;
}

/** `$`? and one to three letters naming a column inside the grid. */
function scanColumn(text: string, i: number): ColumnScan | null {
  let j = i;
  const abs = text.charCodeAt(j) === 0x24;
  if (abs) {
    j += 1;
  }
  const lettersStart = j;
  while (j < text.length && isAsciiLetter(text.charCodeAt(j)) && j - lettersStart < 4) {
    j += 1;
  }
  if (j === lettersStart || j - lettersStart > 3) {
    return null;
  }
  const col = lettersToCol(text.slice(lettersStart, j));
  return col === null ? null : { end: j, col, abs };
}

/** `$`? and a row number 1..1,048,576 (no leading zero). */
function scanRow(text: string, i: number): RowScan | null {
  let j = i;
  const abs = text.charCodeAt(j) === 0x24;
  if (abs) {
    j += 1;
  }
  const first = text.charCodeAt(j);
  if (!(first >= 0x31 && first <= 0x39)) {
    return null;
  }
  const digitsStart = j;
  while (j < text.length && isDigit(text.charCodeAt(j)) && j - digitsStart < 8) {
    j += 1;
  }
  if (isDigit(text.charCodeAt(j))) {
    return null;
  }
  const row = Number(text.slice(digitsStart, j)) - 1;
  return row < MAX_ROWS ? { end: j, row, abs } : null;
}

function scanCell(text: string, i: number): { end: number; part: RefPart } | null {
  const column = scanColumn(text, i);
  if (!column) {
    return null;
  }
  const row = scanRow(text, column.end);
  if (!row) {
    return null;
  }
  return { end: row.end, part: { row: row.row, col: column.col, rowAbs: row.abs, colAbs: column.abs } };
}

/** Whether a reference may end at `j`: no identifier character, `(` or `[`
 * follows it. */
function atBoundary(text: string, j: number): boolean {
  if (j >= text.length) {
    return true;
  }
  const code = text.charCodeAt(j);
  return code !== 0x28 && code !== 0x5b && !isIdentifierPart(text, j);
}

interface AreaScan {
  end: number;
  start: RefPart;
  last?: RefPart;
}

/** A cell, a cell range, a column range or a row range at `i`, ending at a
 * token boundary — the longest that does. */
function scanArea(text: string, i: number): AreaScan | null {
  const first = scanCell(text, i);
  if (first) {
    if (text.charCodeAt(first.end) === 0x3a) {
      const second = scanCell(text, first.end + 1);
      if (second && atBoundary(text, second.end)) {
        return { end: second.end, start: first.part, last: second.part };
      }
    }
    return atBoundary(text, first.end) ? { end: first.end, start: first.part } : null;
  }
  const column = scanColumn(text, i);
  if (column && text.charCodeAt(column.end) === 0x3a) {
    const second = scanColumn(text, column.end + 1);
    if (second && atBoundary(text, second.end)) {
      return {
        end: second.end,
        start: { col: column.col, colAbs: column.abs, rowAbs: false },
        last: { col: second.col, colAbs: second.abs, rowAbs: false },
      };
    }
  }
  const row = scanRow(text, i);
  if (row && text.charCodeAt(row.end) === 0x3a) {
    const second = scanRow(text, row.end + 1);
    if (second && atBoundary(text, second.end)) {
      return {
        end: second.end,
        start: { row: row.row, rowAbs: row.abs, colAbs: false },
        last: { row: second.row, rowAbs: second.abs, colAbs: false },
      };
    }
  }
  return null;
}

function referenceOf(area: AreaScan, sheet?: string, sheetEnd?: string): ReferenceNode {
  const ref: ReferenceNode = { start: area.start };
  if (sheet !== undefined) {
    ref.sheet = sheet;
  }
  if (sheetEnd !== undefined) {
    ref.sheetEnd = sheetEnd;
  }
  if (area.last) {
    ref.end = area.last;
  }
  return ref;
}

// --- quoted names, errors, brackets -------------------------------------------------

/** A `'…'` run at `i` (`''` is one quote): its end (past the closing quote)
 * and unescaped value; null when it never closes. */
export function scanQuoted(text: string, i: number): { end: number; value: string } | null {
  let value = "";
  let j = i + 1;
  while (j < text.length) {
    if (text.charCodeAt(j) === 0x27) {
      if (text.charCodeAt(j + 1) === 0x27) {
        value += "'";
        j += 2;
        continue;
      }
      return { end: j + 1, value };
    }
    value += text[j];
    j += 1;
  }
  return null;
}

/** Whether `word` (upper-case ASCII) is at `i`, case-insensitively. */
function matchesAt(text: string, i: number, word: string): boolean {
  return text.length - i >= word.length && text.slice(i, i + word.length).toUpperCase() === word;
}

/** The error code literal at `i` (`#` …), or null. */
function errorCodeAt(text: string, i: number): string | null {
  for (const code of UPPER_ERROR_CODES) {
    if (matchesAt(text, i, code)) {
      return code;
    }
  }
  return null;
}

/** The end of the balanced `[…]` run at `i` (`'` escapes the next
 * character), or -1 when it never closes. */
function balancedEnd(text: string, i: number): number {
  let depth = 0;
  let j = i;
  while (j < text.length) {
    const code = text.charCodeAt(j);
    if (code === 0x27) {
      j += 2;
      continue;
    }
    if (code === 0x5b) {
      depth += 1;
    } else if (code === 0x5d) {
      depth -= 1;
      if (depth === 0) {
        return j + 1;
      }
    }
    j += 1;
  }
  return -1;
}

/** The end of an OpenFormula bracket reference at `i` (`[` … `]`, quoted
 * sheet names may hold `]`), or -1. */
export function bracketReferenceEnd(text: string, i: number): number {
  let j = i + 1;
  while (j < text.length) {
    const code = text.charCodeAt(j);
    if (code === 0x27) {
      const quoted = scanQuoted(text, j);
      if (!quoted) {
        return -1;
      }
      j = quoted.end;
      continue;
    }
    if (code === 0x5d) {
      return j + 1;
    }
    if (code === 0x5b) {
      return -1;
    }
    j += 1;
  }
  return -1;
}

// --- tokens -----------------------------------------------------------------------

/** What scanning one token at a position found (its text is the source
 * between `start` and `end`). */
interface Scan {
  kind: TokenKind;
  end: number;
  ref?: ReferenceNode;
  unterminated?: true;
}

/** What follows a sheet prefix (`p` is past its `!`): `#REF!`, an area, or
 * a name; null when nothing referential does. */
function scanAfterPrefix(text: string, p: number, sheet: string, sheetEnd?: string): Scan | null {
  if (matchesAt(text, p, "#REF!")) {
    return { kind: "error", end: p + 5 };
  }
  const area = scanArea(text, p);
  if (area) {
    return { kind: "ref", end: area.end, ref: referenceOf(area, sheet, sheetEnd) };
  }
  if (isIdentifierStart(text, p)) {
    const end = identifierEnd(text, p);
    return { kind: text.charCodeAt(end) === 0x28 ? "function" : "name", end };
  }
  return null;
}

/** The end of the reference an external prefix leads into (`p` is past its
 * `!`): an area, `#REF!` or a name; `p` itself when none follows. */
function externalReferenceEnd(text: string, p: number): number {
  if (matchesAt(text, p, "#REF!")) {
    return p + 5;
  }
  const area = scanArea(text, p);
  if (area) {
    return area.end;
  }
  return isIdentifierStart(text, p) ? identifierEnd(text, p) : p;
}

/** A `[`-led token at `i`: an Excel external reference (`[1]Sheet1!A1`,
 * `[1]!Name`), an OpenFormula one kept verbatim (`['file:…'#$S.A1]`), or a
 * structured reference. */
function scanBracket(text: string, i: number): Scan {
  const book = /^\[\d+\]/.exec(text.slice(i, i + 12));
  if (book) {
    let j = i + book[0].length;
    if (isIdentifierStart(text, j)) {
      const first = identifierEnd(text, j);
      if (text.charCodeAt(first) === 0x3a && isIdentifierStart(text, first + 1)) {
        const second = identifierEnd(text, first + 1);
        j = text.charCodeAt(second) === 0x21 ? second : first;
      } else {
        j = first;
      }
    }
    if (text.charCodeAt(j) === 0x21) {
      return { kind: "external", end: Math.max(j + 1, externalReferenceEnd(text, j + 1)) };
    }
  }
  if (text.charCodeAt(i + 1) === 0x27) {
    const quoted = scanQuoted(text, i + 1);
    if (quoted && text.charCodeAt(quoted.end) === 0x23) {
      const end = bracketReferenceEnd(text, i);
      if (end > 0) {
        return { kind: "external", end };
      }
    }
  }
  const end = balancedEnd(text, i);
  return end < 0 ? { kind: "structured", end: text.length, unterminated: true } : { kind: "structured", end };
}

/** A token starting with an identifier at `i`. */
function scanIdentifier(text: string, i: number): Scan {
  const end = identifierEnd(text, i);
  const after = text.charCodeAt(end);
  if (after === 0x21) {
    const prefixed = scanAfterPrefix(text, end + 1, text.slice(i, end));
    return prefixed ?? { kind: "unknown", end: end + 1 };
  }
  if (after === 0x3a && isIdentifierStart(text, end + 1)) {
    const last = identifierEnd(text, end + 1);
    if (text.charCodeAt(last) === 0x21) {
      const prefixed = scanAfterPrefix(text, last + 1, text.slice(i, end), text.slice(end + 1, last));
      return prefixed ?? { kind: "unknown", end: last + 1 };
    }
  }
  if (after === 0x28) {
    return { kind: "function", end };
  }
  const area = scanArea(text, i);
  if (area) {
    return { kind: "ref", end: area.end, ref: referenceOf(area) };
  }
  if (after === 0x5b) {
    const close = balancedEnd(text, end);
    return close < 0 ? { kind: "structured", end: text.length, unterminated: true } : { kind: "structured", end: close };
  }
  const word = text.slice(i, end).toUpperCase();
  return { kind: word === "TRUE" || word === "FALSE" ? "boolean" : "name", end };
}

/** A token starting with `'` at `i`: a quoted sheet prefix and what follows
 * it, or an external one (`'[1]My Sheet'!A1`). */
function scanQuotedPrefix(text: string, i: number): Scan {
  const quoted = scanQuoted(text, i);
  if (!quoted) {
    return { kind: "unknown", end: i + 1 };
  }
  if (text.charCodeAt(quoted.end) !== 0x21) {
    return { kind: "unknown", end: quoted.end };
  }
  if (/^\[\d+\]/.test(quoted.value)) {
    return { kind: "external", end: Math.max(quoted.end + 1, externalReferenceEnd(text, quoted.end + 1)) };
  }
  const colon = quoted.value.indexOf(":");
  const prefixed =
    colon < 0
      ? scanAfterPrefix(text, quoted.end + 1, quoted.value)
      : scanAfterPrefix(text, quoted.end + 1, quoted.value.slice(0, colon), quoted.value.slice(colon + 1));
  return prefixed ?? { kind: "unknown", end: quoted.end + 1 };
}

/** A number at `i`: digits with an optional fraction and exponent, or `.`
 * then digits (`1`, `1.5`, `.5`, `1.`, `1E3`, `1e-3`). */
function numberEnd(text: string, i: number): number {
  let j = i;
  while (isDigit(text.charCodeAt(j))) {
    j += 1;
  }
  if (text.charCodeAt(j) === 0x2e) {
    j += 1;
    while (isDigit(text.charCodeAt(j))) {
      j += 1;
    }
  }
  const e = text.charCodeAt(j);
  if (e === 0x45 || e === 0x65) {
    let k = j + 1;
    const sign = text.charCodeAt(k);
    if (sign === 0x2b || sign === 0x2d) {
      k += 1;
    }
    if (isDigit(text.charCodeAt(k))) {
      while (isDigit(text.charCodeAt(k))) {
        k += 1;
      }
      j = k;
    }
  }
  return j;
}

function startsNumber(text: string, i: number): boolean {
  return isDigit(text.charCodeAt(i)) || (text.charCodeAt(i) === 0x2e && isDigit(text.charCodeAt(i + 1)));
}

/** The bracket a comma or semicolon belongs to. */
type Bracket = "call" | "group" | "brace";

function scanToken(text: string, i: number, stack: readonly Bracket[]): Scan {
  const code = text.charCodeAt(i);
  const top = stack.at(-1);
  switch (code) {
    case 0x22: {
      // A string: "" is one quote.
      let j = i + 1;
      while (j < text.length) {
        if (text.charCodeAt(j) === 0x22) {
          if (text.charCodeAt(j + 1) === 0x22) {
            j += 2;
            continue;
          }
          return { kind: "string", end: j + 1 };
        }
        j += 1;
      }
      return { kind: "string", end: text.length, unterminated: true };
    }
    case 0x27:
      return scanQuotedPrefix(text, i);
    case 0x5b:
      return scanBracket(text, i);
    case 0x23: {
      const error = errorCodeAt(text, i);
      if (error === null) {
        return { kind: "unknown", end: i + 1 };
      }
      if (error === "#REF!") {
        const area = scanArea(text, i + 5);
        if (area) {
          return { kind: "error", end: area.end };
        }
      }
      return { kind: "error", end: i + error.length };
    }
    case 0x24: {
      const area = scanArea(text, i);
      return area ? { kind: "ref", end: area.end, ref: referenceOf(area) } : { kind: "unknown", end: i + 1 };
    }
    case 0x28:
      return { kind: "open", end: i + 1 };
    case 0x29:
      return { kind: "close", end: i + 1 };
    case 0x7b:
      return { kind: "arrayOpen", end: i + 1 };
    case 0x7d:
      return { kind: "arrayClose", end: i + 1 };
    case 0x2c:
      return { kind: top === "brace" ? "arrayColSep" : top === "call" ? "argSep" : "union", end: i + 1 };
    case 0x3b:
      return { kind: top === "brace" ? "arrayRowSep" : "unknown", end: i + 1 };
    case 0x2b:
    case 0x2d:
      if (top === "brace" && startsNumber(text, i + 1)) {
        return { kind: "number", end: numberEnd(text, i + 1) };
      }
      return { kind: "operator", end: i + 1 };
    case 0x2a:
    case 0x2f:
    case 0x5e:
    case 0x26:
    case 0x3d:
    case 0x25:
    case 0x3a:
      return { kind: "operator", end: i + 1 };
    case 0x3c: {
      const next = text.charCodeAt(i + 1);
      return { kind: "operator", end: next === 0x3e || next === 0x3d ? i + 2 : i + 1 };
    }
    case 0x3e:
      return { kind: "operator", end: text.charCodeAt(i + 1) === 0x3d ? i + 2 : i + 1 };
    default:
      break;
  }
  if (isDigit(code)) {
    const area = scanArea(text, i);
    return area ? { kind: "ref", end: area.end, ref: referenceOf(area) } : { kind: "number", end: numberEnd(text, i) };
  }
  if (code === 0x2e && isDigit(text.charCodeAt(i + 1))) {
    return { kind: "number", end: numberEnd(text, i) };
  }
  if (isIdentifierStart(text, i)) {
    return scanIdentifier(text, i);
  }
  return { kind: "unknown", end: i + unitsAt(text, i) };
}

/** Splits canonical formula text into tokens (see the module header). Never
 * throws; the last token is always `end`. */
export function tokenize(text: string): Token[] {
  const tokens: Token[] = [];
  const stack: Bracket[] = [];
  let i = 0;
  for (;;) {
    const leadStart = i;
    while (i < text.length && isWhitespace(text.charCodeAt(i))) {
      i += 1;
    }
    const lead = text.slice(leadStart, i);
    if (i >= text.length) {
      tokens.push({ kind: "end", text: "", lead, start: text.length });
      return tokens;
    }
    const previous = tokens.at(-1);
    const scan = scanToken(text, i, stack);
    const token: Token = { kind: scan.kind, text: text.slice(i, scan.end), lead, start: i };
    if (scan.ref) {
      token.ref = scan.ref;
    }
    if (scan.unterminated) {
      token.unterminated = true;
    }
    switch (scan.kind) {
      case "open":
        stack.push(previous?.kind === "function" && lead === "" ? "call" : "group");
        break;
      case "close":
        if (stack.at(-1) === "call" || stack.at(-1) === "group") {
          stack.pop();
        }
        break;
      case "arrayOpen":
        stack.push("brace");
        break;
      case "arrayClose":
        if (stack.at(-1) === "brace") {
          stack.pop();
        }
        break;
      default:
        break;
    }
    const space = lead.lastIndexOf(" ");
    if (space >= 0 && previous && REFERENCE_ENDING.has(previous.kind) && REFERENCE_STARTING.has(scan.kind)) {
      tokens.push({ kind: "operator", text: " ", lead: lead.slice(0, space), start: leadStart + space });
      token.lead = lead.slice(space + 1);
    }
    tokens.push(token);
    i = scan.end;
  }
}

/** The source text of a token stream: every token's lead and text, in order. */
export function joinTokens(tokens: readonly Token[]): string {
  let out = "";
  for (const token of tokens) {
    out += token.lead + token.text;
  }
  return out;
}

/** The sheet prefix of a reference-bearing token's text as written —
 * `Sheet2!`, `'My Sheet'!`, `Sheet1:Sheet3!` — or "" when it has none. */
export function sheetPrefixOf(text: string): string {
  if (text.charCodeAt(0) === 0x27) {
    const quoted = scanQuoted(text, 0);
    return quoted && text.charCodeAt(quoted.end) === 0x21 ? text.slice(0, quoted.end + 1) : "";
  }
  if (!isIdentifierStart(text, 0)) {
    return "";
  }
  let end = identifierEnd(text, 0);
  if (text.charCodeAt(end) === 0x3a && isIdentifierStart(text, end + 1)) {
    const last = identifierEnd(text, end + 1);
    if (text.charCodeAt(last) === 0x21) {
      end = last;
    }
  }
  return text.charCodeAt(end) === 0x21 ? text.slice(0, end + 1) : "";
}

/** A sheet prefix as written (`'O''Brien'!`, `Jan:Mar!`) → its sheet name(s),
 * unquoted and un-escaped. */
export function parseSheetPrefix(prefix: string): { sheet: string; sheetEnd?: string } {
  const body = prefix.endsWith("!") ? prefix.slice(0, -1) : prefix;
  const unquoted = body.startsWith("'") ? (scanQuoted(body, 0)?.value ?? body) : body;
  const colon = unquoted.indexOf(":");
  return colon < 0 ? { sheet: unquoted } : { sheet: unquoted.slice(0, colon), sheetEnd: unquoted.slice(colon + 1) };
}
