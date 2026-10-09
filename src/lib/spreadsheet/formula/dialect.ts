// Formula text between canonical form and each format's wire form (SPEC-119
// "xlsx dialect", "ODS dialect", "Crossing formats"). Every translation is
// token-level and carries each token's lead, so whitespace survives both
// directions.
//
// xlsx:
// - fromXlsxFormula(wire): one stray leading "=" dropped; every function
//   name stripped of `_xlfn._xlws.` / `_xlfn.` and upper-cased; every name
//   stripped of `_xlpm.` (a LET/LAMBDA parameter).
// - toXlsxFormula(canonical): function names through the table's prefixes
//   (functionNames.ts); `_xlpm.` on every name that is a parameter of an
//   enclosing LET (arguments 1, 3, 5 … but the last) or LAMBDA (every
//   argument but the last), or a use of one inside that call's argument span
//   (lexical scope — any enclosing declaration makes it a parameter). No "=".
//
// ODS — fromOdsFormula(wire, namespaces) takes the whole `table:formula`
// value and the prefix → URI map in scope; the dialect is the prefix's URI
// (an undeclared of:/oooc:/msoxl: prefix falls back to its usual URI):
// - OpenFormula (`of:`, and unprefixed text) and the legacy `oooc:` dialect
//   (the same bracket grammar and `;` separators): `[.A1]` → `A1`,
//   `[.A1:.B2]` → `A1:B2`, `[$Sheet1.A1]`/`[Sheet1.A1]` → `Sheet1!A1`,
//   `[$Sheet1.A1:.B2]`/`[$Sheet1.A1:$Sheet1.B2]` → `Sheet1!A1:B2`,
//   `['My Sheet'.A1]` → `'My Sheet'!A1`, `[$Sheet1.A1:$Sheet3.B2]` →
//   `Sheet1:Sheet3!A1:B2`, `[.A:.C]`/`[.1:.3]` → `A:C`/`1:3`, any bracket
//   reference holding `#REF!` → `#REF!`, `['file:///…'#$Sheet1.A1]` → an
//   `external` token, verbatim; `;` → `,`; `a ~ b` → `(a,b)` — parentheses
//   added unless the union already is a grouping paren's whole content;
//   `a ! b` → `a b`; `{1;2|3;4}` → `{1,2;3,4}`; function names through
//   canonicalFromOds. A bracket reference the grammar cannot read, or an
//   unknown dialect, is a FormulaSyntaxError (the reader then keeps the wire
//   text, formulaUnparsed).
// - `msoxl:` (Excel's own syntax): "=" stripped, then fromXlsxFormula.
// - toOdsFormula(canonical) inverts it with an `of:=` prefix: `[.A1]` for a
//   same-sheet reference, `[$Sheet.A1:.B2]` for another sheet (absolute
//   sheet, LibreOffice's default), sheet names quoted by print.ts's
//   predicate, `,` → `;`, union → `~` (Excel's parentheses kept — harmless
//   in OpenFormula), intersection → `!`, arrays → `{…;…|…}`.
// Known non-identities (toOdsFormula(fromOdsFormula(w)) !== w): a relative-
// sheet `[Sheet2.A1]` comes back as `[$Sheet2.A1]`; a same-sheet second
// endpoint `[$S.A1:$S.B2]` as `[$S.A1:.B2]`; a `~` union gains its
// parentheses; a `#REF!` reference comes back as the bare error. Only dirty
// sheets (SPEC-122) ever take this path — a clean sheet is spliced verbatim.
//
// formulaForTarget(cell, target) is the one rule for a cell crossing to the
// other format (SPEC-121's "Download a copy as…"): a parsed formula crosses
// unchanged (it is canonical); an unparsed formula or a data table cannot,
// so the cell drops to its cached value and reports the drop; a dynamic
// array crossing to .ods keeps formula + arrayRange (a matrix formula) and
// loses `cm` and `dynamicArray`.
import { colToLetters, lettersToCol, type CellRef } from "../address";
import { MAX_ROWS, type Cell, type SpreadsheetFormat } from "../model";
import { canonicalFromOds, canonicalFromXlsx, odsFromCanonical, xlsxFromCanonical } from "./functionNames";
import { FormulaSyntaxError, parseFormulaAt } from "./parse";
import { formatReference, sheetNameNeedsQuotes } from "./print";
import type { RefPart, ReferenceNode } from "./references";
import {
  bracketReferenceEnd,
  identifierEnd,
  isIdentifierStart,
  joinTokens,
  scanQuoted,
  sheetPrefixOf,
  tokenize,
  type Token,
  type TokenKind,
} from "./tokens";

export const OPENFORMULA_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
export const OOOC_NAMESPACE = "http://openoffice.org/2004/calc";
export const MSOXL_NAMESPACE = "http://schemas.microsoft.com/office/excel/formula";

/** The URI an undeclared dialect prefix stands for. */
const USUAL_PREFIXES: ReadonlyMap<string, string> = new Map([
  ["of", OPENFORMULA_NAMESPACE],
  ["oooc", OOOC_NAMESPACE],
  ["msoxl", MSOXL_NAMESPACE],
]);

// --- xlsx -------------------------------------------------------------------------

const XLPM = "_xlpm.";

/** A function token's text with its name (not a sheet prefix) mapped. */
function mapFunctionName(text: string, map: (name: string) => string): string {
  const prefix = sheetPrefixOf(text);
  return prefix + map(text.slice(prefix.length));
}

/** xlsx `<f>` text → canonical (see the module header). */
export function fromXlsxFormula(wire: string): string {
  const tokens = tokenize(wire.startsWith("=") ? wire.slice(1) : wire);
  for (const token of tokens) {
    if (token.kind === "function") {
      token.text = mapFunctionName(token.text, canonicalFromXlsx);
    } else if (token.kind === "name" && token.text.slice(0, XLPM.length).toLowerCase() === XLPM) {
      token.text = token.text.slice(XLPM.length);
    }
  }
  return joinTokens(tokens);
}

/** A LET or LAMBDA call's span (token indices of its `(` and `)`) and the
 * parameter names it declares (upper-case), with their token indices. */
interface ParameterScope {
  from: number;
  to: number;
  names: Set<string>;
  declarations: Set<number>;
}

/** Every LET/LAMBDA call of a canonical token stream and its parameters. */
function parameterScopes(tokens: readonly Token[]): ParameterScope[] {
  const scopes: ParameterScope[] = [];
  const stack: Array<{ name: string | null; from: number; argStart: number; args: Array<[number, number]> }> = [];
  tokens.forEach((token, index) => {
    if (token.kind === "open") {
      const previous = tokens[index - 1];
      const name = previous?.kind === "function" ? canonicalFromXlsx(previous.text.slice(sheetPrefixOf(previous.text).length)) : null;
      stack.push({ name, from: index, argStart: index + 1, args: [] });
    } else if (token.kind === "argSep") {
      const top = stack.at(-1);
      if (top && top.name !== null) {
        top.args.push([top.argStart, index - 1]);
        top.argStart = index + 1;
      }
    } else if (token.kind === "close") {
      const top = stack.pop();
      if (!top || (top.name !== "LET" && top.name !== "LAMBDA")) {
        return;
      }
      if (!(top.args.length === 0 && top.argStart === index)) {
        top.args.push([top.argStart, index - 1]);
      }
      const scope: ParameterScope = { from: top.from, to: index, names: new Set(), declarations: new Set() };
      top.args.forEach(([first, last], position) => {
        const declares = position < top.args.length - 1 && (top.name === "LAMBDA" || position % 2 === 0);
        const only = tokens[first];
        if (declares && first === last && only?.kind === "name" && sheetPrefixOf(only.text) === "") {
          scope.names.add(only.text.toUpperCase());
          scope.declarations.add(first);
        }
      });
      scopes.push(scope);
    }
  });
  return scopes;
}

/** canonical → xlsx `<f>` text (see the module header). */
export function toXlsxFormula(canonical: string): string {
  const tokens = tokenize(canonical);
  const scopes = parameterScopes(tokens);
  tokens.forEach((token, index) => {
    if (token.kind === "function") {
      token.text = mapFunctionName(token.text, xlsxFromCanonical);
    } else if (token.kind === "name" && sheetPrefixOf(token.text) === "") {
      const name = token.text.toUpperCase();
      const parameter = scopes.some(
        (scope) => scope.declarations.has(index) || (scope.from < index && index < scope.to && scope.names.has(name)),
      );
      if (parameter) {
        token.text = XLPM + token.text;
      }
    }
  });
  return joinTokens(tokens);
}

// --- OpenFormula ------------------------------------------------------------------------

function isWhitespace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

function isDigit(code: number): boolean {
  return code >= 0x30 && code <= 0x39;
}

/** The end of a number at `i` (the canonical grammar's). */
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
    if (text.charCodeAt(k) === 0x2b || text.charCodeAt(k) === 0x2d) {
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

const OPENFORMULA_ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];

/** OpenFormula text (after its "=") → tokens of the same shape as
 * tokens.ts's: bracket references are `ref` tokens (their text raw), `;` is
 * `argSep`/`arrayColSep`, `|` `arrayRowSep`, `~` `union`, `!` the
 * intersection `operator`. Total, like tokenize. */
export function tokenizeOpenFormula(text: string): Token[] {
  const tokens: Token[] = [];
  const stack: Array<"call" | "group" | "brace"> = [];
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
    const code = text.charCodeAt(i);
    const top = stack.at(-1);
    let kind: TokenKind = "unknown";
    let end = i + 1;
    if (code === 0x5b) {
      const close = bracketReferenceEnd(text, i);
      if (close > 0) {
        kind = "ref";
        end = close;
      }
    } else if (code === 0x22) {
      kind = "string";
      end = text.length;
      for (let j = i + 1; j < text.length; j += 1) {
        if (text.charCodeAt(j) === 0x22) {
          if (text.charCodeAt(j + 1) === 0x22) {
            j += 1;
            continue;
          }
          end = j + 1;
          break;
        }
      }
    } else if (code === 0x23) {
      const error = OPENFORMULA_ERRORS.find((candidate) => text.slice(i, i + candidate.length).toUpperCase() === candidate);
      if (error) {
        kind = "error";
        end = i + error.length;
      }
    } else if (code === 0x28) {
      kind = "open";
      stack.push(tokens.at(-1)?.kind === "function" && lead === "" ? "call" : "group");
    } else if (code === 0x29) {
      kind = "close";
      if (top === "call" || top === "group") {
        stack.pop();
      }
    } else if (code === 0x7b) {
      kind = "arrayOpen";
      stack.push("brace");
    } else if (code === 0x7d) {
      kind = "arrayClose";
      if (top === "brace") {
        stack.pop();
      }
    } else if (code === 0x3b) {
      kind = top === "brace" ? "arrayColSep" : top === "call" ? "argSep" : "unknown";
    } else if (code === 0x7c) {
      kind = top === "brace" ? "arrayRowSep" : "unknown";
    } else if (code === 0x7e) {
      kind = "union";
    } else if ("!+-*/^&=%:".includes(text[i]!)) {
      kind = "operator";
    } else if (code === 0x3c) {
      kind = "operator";
      end = text.charCodeAt(i + 1) === 0x3e || text.charCodeAt(i + 1) === 0x3d ? i + 2 : i + 1;
    } else if (code === 0x3e) {
      kind = "operator";
      end = text.charCodeAt(i + 1) === 0x3d ? i + 2 : i + 1;
    } else if (isDigit(code) || (code === 0x2e && isDigit(text.charCodeAt(i + 1)))) {
      kind = "number";
      end = numberEnd(text, i);
    } else if (isIdentifierStart(text, i)) {
      end = identifierEnd(text, i);
      kind = text.charCodeAt(end) === 0x28 ? "function" : "name";
    } else if (code >= 0xd800 && code <= 0xdbff && i + 1 < text.length) {
      end = i + 2;
    }
    tokens.push({ kind, text: text.slice(i, end), lead, start: i });
    i = end;
  }
}

interface OdfAddress {
  part: RefPart;
  type: "cell" | "column" | "row";
}

const ODF_ADDRESS = /^(\$?)([A-Za-z]{1,3})?(\$?)([0-9]{1,7})?/;

/** One OpenFormula endpoint at `i` of a bracket's content: an optional
 * (`$`-absolute, possibly quoted) sheet, `.`, and a cell, column or row. */
function odfEndpoint(inner: string, i: number): { end: number; sheet?: string; address: OdfAddress } | null {
  let j = i;
  let sheet: string | undefined;
  if (inner[j] !== ".") {
    if (inner[j] === "$") {
      j += 1;
    }
    if (inner[j] === "'") {
      const quoted = scanQuoted(inner, j);
      if (!quoted) {
        return null;
      }
      sheet = quoted.value;
      j = quoted.end;
    } else {
      const start = j;
      while (j < inner.length && inner[j] !== "." && inner[j] !== ":") {
        j += 1;
      }
      sheet = inner.slice(start, j);
    }
    if (sheet === "" || inner[j] !== ".") {
      return null;
    }
  }
  const match = ODF_ADDRESS.exec(inner.slice(j + 1));
  if (!match || (match[2] === undefined && match[4] === undefined)) {
    return null;
  }
  const end = j + 1 + match[0].length;
  const col = match[2] === undefined ? undefined : lettersToCol(match[2]);
  const row = match[4] === undefined ? undefined : Number(match[4]) - 1;
  if (col === null || (row !== undefined && (row < 0 || row >= MAX_ROWS))) {
    return null;
  }
  // A row-only part's `$` is captured by the first group.
  const part: RefPart =
    col === undefined ? { rowAbs: match[1] === "$" || match[3] === "$", colAbs: false } : { rowAbs: match[3] === "$", colAbs: match[1] === "$" };
  if (col !== undefined) {
    part.col = col;
  }
  if (row !== undefined) {
    part.row = row;
  }
  const type = col !== undefined && row !== undefined ? "cell" : col !== undefined ? "column" : "row";
  return sheet === undefined ? { end, address: { part, type } } : { end, sheet, address: { part, type } };
}

/** A bracket reference's content (`.A1:.B2`, `$'My Sheet'.A1`) → a
 * ReferenceNode; null when the grammar cannot read it. */
function odfReference(inner: string): ReferenceNode | null {
  const first = odfEndpoint(inner, 0);
  if (!first) {
    return null;
  }
  const ref: ReferenceNode = { start: first.address.part };
  if (first.sheet !== undefined) {
    ref.sheet = first.sheet;
  }
  if (first.end === inner.length) {
    return first.address.type === "cell" ? ref : null;
  }
  if (inner[first.end] !== ":") {
    return null;
  }
  const second = odfEndpoint(inner, first.end + 1);
  if (!second || second.end !== inner.length || second.address.type !== first.address.type) {
    return null;
  }
  if (second.sheet !== undefined) {
    if (first.sheet === undefined) {
      return null;
    }
    if (second.sheet !== first.sheet) {
      ref.sheetEnd = second.sheet;
    }
  }
  const [a, b] = [first.address.part, second.address.part];
  const sameCell = a.row === b.row && a.col === b.col && a.rowAbs === b.rowAbs && a.colAbs === b.colAbs;
  // A 3-D single cell can only be written as a range in OpenFormula.
  if (!(ref.sheetEnd !== undefined && first.address.type === "cell" && sameCell)) {
    ref.end = b;
  }
  return ref;
}

/** Whether a bracket reference is an external one (`['file:…'#…]`). */
function isExternalBracket(text: string): boolean {
  const inner = text.slice(1, -1);
  const at = inner.startsWith("$'") ? 1 : 0;
  const quoted = inner[at] === "'" ? scanQuoted(inner, at) : null;
  return quoted !== null && inner[quoted.end] === "#";
}

/** Matching bracket indices of a token stream (open ↔ close, both ways). */
function matchingBrackets(tokens: readonly Token[]): Map<number, number> {
  const pairs = new Map<number, number>();
  const stack: number[] = [];
  tokens.forEach((token, index) => {
    if (token.kind === "open" || token.kind === "arrayOpen") {
      stack.push(index);
    } else if (token.kind === "close" || token.kind === "arrayClose") {
      const open = stack.pop();
      if (open !== undefined) {
        pairs.set(open, index);
        pairs.set(index, open);
      }
    }
  });
  return pairs;
}

/** `~`, `:` or `!` — the operators that bind at least as tightly as union. */
function isChainOperator(token: Token | undefined): boolean {
  return token !== undefined && (token.kind === "union" || (token.kind === "operator" && (token.text === ":" || token.text === "!")));
}

/** Each maximal `~` chain of an OpenFormula token stream that needs the
 * parentheses Excel's union requires: [first, last] token indices. A chain's
 * operands are primaries (a reference, name, call or bracketed group) joined
 * by `:`, `!` and `~` — the operators binding at least as tightly as union. */
function unionsNeedingParens(tokens: readonly Token[]): Array<[number, number]> {
  const pairs = matchingBrackets(tokens);
  const operandStart = (at: number): number => {
    const token = tokens[at];
    if (token?.kind !== "close" && token?.kind !== "arrayClose") {
      return at;
    }
    const open = pairs.get(at) ?? at;
    return tokens[open - 1]?.kind === "function" && token.kind === "close" ? open - 1 : open;
  };
  const operandEnd = (at: number): number => {
    const token = tokens[at];
    if (token?.kind === "function") {
      return pairs.get(at + 1) ?? at;
    }
    return token?.kind === "open" || token?.kind === "arrayOpen" ? (pairs.get(at) ?? at) : at;
  };
  const chains = new Map<string, [number, number]>();
  tokens.forEach((token, index) => {
    if (token.kind !== "union" || index === 0) {
      return;
    }
    let first = operandStart(index - 1);
    while (first >= 2 && isChainOperator(tokens[first - 1])) {
      first = operandStart(first - 2);
    }
    let last = operandEnd(index + 1);
    while (last + 2 < tokens.length && isChainOperator(tokens[last + 1])) {
      last = operandEnd(last + 2);
    }
    chains.set(`${first}:${last}`, [first, last]);
  });
  return [...chains.values()].filter(([first, last]) => {
    const before = tokens[first - 1];
    const grouped = before?.kind === "open" && tokens[first - 2]?.kind !== "function" && pairs.get(first - 1) === last + 1;
    return !grouped;
  });
}

/** An OpenFormula token's canonical text. */
function canonicalTokenText(token: Token): string {
  switch (token.kind) {
    case "ref": {
      if (isExternalBracket(token.text)) {
        return token.text;
      }
      const inner = token.text.slice(1, -1);
      if (inner.toUpperCase().includes("#REF!")) {
        return "#REF!";
      }
      const ref = odfReference(inner);
      if (!ref) {
        throw new FormulaSyntaxError(`"${token.text}" is not an OpenFormula reference`, token.start);
      }
      return formatReference(ref);
    }
    case "function":
      return canonicalFromOds(token.text);
    case "argSep":
    case "arrayColSep":
    case "union":
      return ",";
    case "arrayRowSep":
      return ";";
    case "operator":
      return token.text === "!" ? " " : token.text;
    default:
      return token.text;
  }
}

/** OpenFormula text (after its "=") → canonical. */
function fromOpenFormula(body: string): string {
  const tokens = tokenizeOpenFormula(body);
  const opens = new Map<number, number>();
  const closes = new Map<number, number>();
  for (const [first, last] of unionsNeedingParens(tokens)) {
    opens.set(first, (opens.get(first) ?? 0) + 1);
    closes.set(last, (closes.get(last) ?? 0) + 1);
  }
  let out = "";
  tokens.forEach((token, index) => {
    out += token.lead + "(".repeat(opens.get(index) ?? 0) + canonicalTokenText(token) + ")".repeat(closes.get(index) ?? 0);
  });
  return out;
}

const DIALECT_PREFIX = /^([A-Za-z_][\w.-]*):/;

/** An ODS `table:formula` value → canonical (see the module header). Throws
 * FormulaSyntaxError for an unknown dialect or an unreadable reference. */
export function fromOdsFormula(wire: string, namespaces: ReadonlyMap<string, string> = new Map()): string {
  const match = wire.startsWith("=") ? null : DIALECT_PREFIX.exec(wire);
  const prefix = match?.[1];
  const body = match ? wire.slice(match[0].length) : wire;
  const formula = body.startsWith("=") ? body.slice(1) : body;
  const uri = prefix === undefined ? OPENFORMULA_NAMESPACE : (namespaces.get(prefix) ?? USUAL_PREFIXES.get(prefix));
  switch (uri) {
    case OPENFORMULA_NAMESPACE:
    case OOOC_NAMESPACE:
      return fromOpenFormula(formula);
    case MSOXL_NAMESPACE:
      return fromXlsxFormula(formula);
    default:
      throw new FormulaSyntaxError(`"${prefix}:" is not a formula language this reader knows`, 0);
  }
}

/** A sheet name as an OpenFormula reference writes it: `$` (absolute), then
 * bare or quoted by the shared predicate. */
function odsSheet(name: string): string {
  return `$${sheetNameNeedsQuotes(name) ? `'${name.replaceAll("'", "''")}'` : name}`;
}

function odsPart(part: RefPart): string {
  let text = "";
  if (part.col !== undefined) {
    text += `${part.colAbs ? "$" : ""}${colToLetters(part.col)}`;
  }
  if (part.row !== undefined) {
    text += `${part.rowAbs ? "$" : ""}${part.row + 1}`;
  }
  return text;
}

/** A reference as an OpenFormula bracket reference — a whole column or row,
 * and a 3-D single cell, as a range (`[$Jan.B5:$Mar.B5]`). */
function odsReference(ref: ReferenceNode): string {
  const sheet = ref.sheet === undefined ? "" : odsSheet(ref.sheet);
  const whole = ref.start.row === undefined || ref.start.col === undefined;
  const end = ref.end ?? (whole || ref.sheetEnd !== undefined ? ref.start : undefined);
  let text = `${sheet}.${odsPart(ref.start)}`;
  if (end) {
    text += `:${ref.sheetEnd === undefined ? "" : odsSheet(ref.sheetEnd)}.${odsPart(end)}`;
  }
  return `[${text}]`;
}

/** A canonical token's OpenFormula text. */
function odsTokenText(token: Token): string {
  switch (token.kind) {
    case "ref":
      return token.ref ? odsReference(token.ref) : token.text;
    case "function":
      return mapFunctionName(token.text, odsFromCanonical);
    case "argSep":
    case "arrayColSep":
      return ";";
    case "arrayRowSep":
      return "|";
    case "union":
      return "~";
    case "operator":
      return token.text === " " ? "!" : token.text;
    case "error":
      return token.text.toUpperCase().includes("#REF!") ? "#REF!" : token.text;
    default:
      return token.text;
  }
}

/** canonical → an ODS `table:formula` value, `of:=` prefixed (see the
 * module header). */
export function toOdsFormula(canonical: string): string {
  let out = "of:=";
  for (const token of tokenize(canonical)) {
    out += token.lead + odsTokenText(token);
  }
  return out;
}

// --- reading and crossing -------------------------------------------------------------

/** A formula as the model stores it: canonical, or the wire text when it
 * does not parse. */
export interface FormulaFields {
  formula: string;
  formulaUnparsed?: true;
}

/** The parse check both readers make: canonical when `translate` succeeds
 * and the result parses (through parseFormulaAt, so a shared block costs one
 * parse), else the wire text verbatim. */
function checkedFormula(wire: string, translate: () => string, host: CellRef): FormulaFields {
  let canonical: string;
  try {
    canonical = translate();
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return { formula: wire, formulaUnparsed: true };
    }
    throw error;
  }
  return parseFormulaAt(canonical, host) instanceof FormulaSyntaxError ? { formula: wire, formulaUnparsed: true } : { formula: canonical };
}

/** An xlsx `<f>` text read at `host`. */
export function readXlsxFormula(wire: string, host: CellRef): FormulaFields {
  return checkedFormula(wire, () => fromXlsxFormula(wire), host);
}

/** An ODS `table:formula` value read at `host`. */
export function readOdsFormula(wire: string, namespaces: ReadonlyMap<string, string>, host: CellRef): FormulaFields {
  return checkedFormula(wire, () => fromOdsFormula(wire, namespaces), host);
}

/** What a crossing drops (SPEC-121 lists "unparsed" and "dataTable" in its
 * "Not included in the … copy" toast). */
export type FormulaDrop = "unparsed" | "dataTable" | "dynamicArray";

/** A cell as it crosses to `target`, the other format (see the module
 * header). The cell is never mutated; its value is always kept. */
export function formulaForTarget(cell: Cell, target: SpreadsheetFormat): { cell: Cell; dropped?: FormulaDrop } {
  if (cell.formulaUnparsed || cell.dataTableXml !== undefined) {
    const { formula: _formula, formulaUnparsed: _unparsed, arrayRange: _range, dynamicArray: _dynamic, cm: _cm, dataTableXml, ...rest } = cell;
    return { cell: rest, dropped: dataTableXml !== undefined ? "dataTable" : "unparsed" };
  }
  if (target === "ods" && (cell.dynamicArray || cell.cm !== undefined)) {
    const { dynamicArray, cm: _cm, ...rest } = cell;
    return dynamicArray ? { cell: rest, dropped: "dynamicArray" } : { cell: rest };
  }
  return { cell };
}
