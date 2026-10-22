// The formula parser (SPEC-119 "The AST", "Precedence", "Literals") —
// recursive descent over tokens.ts's stream, building ast.ts's FormulaNode.
//
// Precedence, tightest first, every binary operator left-associative (so
// `2^3^2` is `(2^3)^2` and `A1=B1=C1` is `(A1=B1)=C1`):
//   `:` (operands must be references) > intersection (a space) > union (a
//   `union` comma — top level or a grouping paren) > prefix `-`/`+` > postfix
//   `%` > `^` > `*` `/` > `+` `-` > `&` > `=` `<>` `<` `<=` `>` `>=`.
// So `-2^2` is `(-2)^2`, `2^-2` is `2^(-2)`, `2^3%` is `2^(3%)` and `-A1:A3`
// is `-(A1:A3)`; a binary operator's right operand may begin with any number
// of prefix operators (`1+-+-1`).
//
// - Literals: a number's value is Number(text) — one overflowing a double
//   (`1E400`) is a syntax error; strings unescape `""`, and an unterminated
//   one is a syntax error; TRUE/FALSE in any case; the fifteen error codes
//   (`Sheet2!#REF!` and `#REF!A1` keep that spelling in `text`); array
//   constants hold signed numbers, strings, booleans and errors only, in
//   rows of equal length.
// - A function call keeps empty argument slots as `missing` nodes:
//   `IF(A1,,1)` has three arguments, `F()` none, `IFERROR(1/0,)` two.
// - FormulaSyntaxError's `offset` is the UTF-16 index of the offending
//   token. Nesting — parentheses, calls, prefix operators — deeper than
//   MAX_PARSE_DEPTH is a FormulaSyntaxError, never a stack overflow.
// - parseFormulaAt caches ASTs in an LRU of FORMULA_AST_CACHE_SIZE entries
//   keyed by relativeFormulaKey (references.ts), so every cell of a shared
//   formula — 100,000 of them — shares ONE AST. `origin` is the host the AST
//   was first parsed for; a later caller resolves a relative part as
//   `part + (host − origin)` (SPEC-120). Cached nodes are frozen: never
//   mutated. (An entry also records its token kinds: two texts whose
//   trivia-free spellings coincide but tokenize differently — `1 2` and `12`
//   — never share an AST.)
import type { CellRef } from "../address";
import { ERROR_CODES, type ErrorCode } from "../model";
import type { ArrayElement, BinaryOp, FormulaNode, UnaryOp } from "./ast";
import type { ReferenceNode } from "./references";
import { relativeKeyOf } from "./references";
import { parseSheetPrefix, sheetPrefixOf, tokenize, type Token } from "./tokens";

export const MAX_PARSE_DEPTH = 256;
export const FORMULA_AST_CACHE_SIZE = 50_000;

/** A formula the grammar does not accept. `offset` is the UTF-16 index of
 * the offending token in the parsed text. */
export class FormulaSyntaxError extends Error {
  readonly offset: number;

  constructor(message: string, offset: number) {
    super(message);
    this.name = "FormulaSyntaxError";
    this.offset = offset;
  }
}

const COMPARISONS: ReadonlySet<string> = new Set(["=", "<>", "<", "<=", ">", ">="]);
const CODE_BY_UPPER: ReadonlyMap<string, ErrorCode> = new Map(ERROR_CODES.map((code) => [code.toUpperCase(), code]));

/** The node kinds `:` accepts as operands — what can evaluate to a
 * reference. */
const REFERENCE_KINDS: ReadonlySet<FormulaNode["kind"]> = new Set(["ref", "name", "function", "paren", "structured", "external", "error"]);
const REFERENCE_OPS: ReadonlySet<BinaryOp> = new Set([":", " ", ","]);

/** The error code an error token spells: its `#…` part, any case. */
function errorCodeOf(text: string): ErrorCode | undefined {
  const body = text.slice(sheetPrefixOf(text).length).toUpperCase();
  for (const [upper, code] of CODE_BY_UPPER) {
    if (body.startsWith(upper)) {
      return code;
    }
  }
  return undefined;
}

class Parser {
  private index = 0;
  private depth = 0;

  constructor(private readonly tokens: readonly Token[]) {}

  parse(): FormulaNode {
    const node = this.expression();
    const next = this.peek();
    if (next.kind !== "end") {
      this.fail(next);
    }
    return node;
  }

  private peek(): Token {
    return this.tokens[this.index]!;
  }

  private fail(token: Token, message = `Unexpected ${token.kind === "end" ? "end of formula" : `"${token.text}"`}`): never {
    throw new FormulaSyntaxError(message, token.start);
  }

  private enter(token: Token): void {
    this.depth += 1;
    if (this.depth > MAX_PARSE_DEPTH) {
      this.fail(token, `Formula nested deeper than ${MAX_PARSE_DEPTH} levels`);
    }
  }

  private leave(): void {
    this.depth -= 1;
  }

  private isOperator(token: Token, ...texts: string[]): boolean {
    return token.kind === "operator" && texts.includes(token.text);
  }

  private binary(op: BinaryOp, left: FormulaNode, right: FormulaNode): FormulaNode {
    return { kind: "binary", op, left, right, from: left.from, to: right.to };
  }

  private expression(): FormulaNode {
    let left = this.concatenation();
    while (this.peek().kind === "operator" && COMPARISONS.has(this.peek().text)) {
      const op = this.peek().text as BinaryOp;
      this.index += 1;
      left = this.binary(op, left, this.concatenation());
    }
    return left;
  }

  private concatenation(): FormulaNode {
    let left = this.additive();
    while (this.isOperator(this.peek(), "&")) {
      this.index += 1;
      left = this.binary("&", left, this.additive());
    }
    return left;
  }

  private additive(): FormulaNode {
    let left = this.multiplicative();
    while (this.isOperator(this.peek(), "+", "-")) {
      const op = this.peek().text as BinaryOp;
      this.index += 1;
      left = this.binary(op, left, this.multiplicative());
    }
    return left;
  }

  private multiplicative(): FormulaNode {
    let left = this.exponent();
    while (this.isOperator(this.peek(), "*", "/")) {
      const op = this.peek().text as BinaryOp;
      this.index += 1;
      left = this.binary(op, left, this.exponent());
    }
    return left;
  }

  private exponent(): FormulaNode {
    let left = this.percent();
    while (this.isOperator(this.peek(), "^")) {
      this.index += 1;
      left = this.binary("^", left, this.percent());
    }
    return left;
  }

  private percent(): FormulaNode {
    let operand = this.prefix();
    while (this.isOperator(this.peek(), "%")) {
      operand = { kind: "unary", op: "%", operand, from: operand.from, to: this.index };
      this.index += 1;
    }
    return operand;
  }

  private prefix(): FormulaNode {
    const token = this.peek();
    if (!this.isOperator(token, "-", "+")) {
      return this.union();
    }
    const from = this.index;
    this.enter(token);
    this.index += 1;
    const operand = this.prefix();
    this.leave();
    return { kind: "unary", op: token.text as UnaryOp, operand, from, to: operand.to };
  }

  private union(): FormulaNode {
    let left = this.intersection();
    while (this.peek().kind === "union") {
      this.index += 1;
      left = this.binary(",", left, this.intersection());
    }
    return left;
  }

  private intersection(): FormulaNode {
    let left = this.range();
    while (this.isOperator(this.peek(), " ")) {
      this.index += 1;
      left = this.binary(" ", left, this.range());
    }
    return left;
  }

  /** A `:` operand must be something that can be a reference. */
  private requireReference(node: FormulaNode): void {
    const reference = node.kind === "binary" ? REFERENCE_OPS.has(node.op) : REFERENCE_KINDS.has(node.kind);
    if (!reference || (node.kind === "error" && node.code !== "#REF!")) {
      this.fail(this.tokens[node.from]!, "A range needs references on both sides of \":\"");
    }
  }

  private range(): FormulaNode {
    let left = this.primary();
    while (this.isOperator(this.peek(), ":")) {
      this.requireReference(left);
      this.index += 1;
      const right = this.primary();
      this.requireReference(right);
      left = this.binary(":", left, right);
    }
    return left;
  }

  private primary(): FormulaNode {
    const token = this.peek();
    const at = this.index;
    switch (token.kind) {
      case "number": {
        this.index += 1;
        return { kind: "number", value: this.numberValue(token), from: at, to: at };
      }
      case "string": {
        if (token.unterminated) {
          this.fail(token, "Unterminated string");
        }
        this.index += 1;
        return { kind: "string", value: token.text.slice(1, -1).replaceAll('""', '"'), from: at, to: at };
      }
      case "boolean":
        this.index += 1;
        return { kind: "boolean", value: token.text.toUpperCase() === "TRUE", from: at, to: at };
      case "error": {
        const code = errorCodeOf(token.text);
        if (code === undefined) {
          this.fail(token);
        }
        this.index += 1;
        return { kind: "error", code, text: token.text, from: at, to: at };
      }
      case "ref":
        this.index += 1;
        return { kind: "ref", ref: token.ref as ReferenceNode, from: at, to: at };
      case "name":
        this.index += 1;
        return this.name(token, at);
      case "function":
        return this.call();
      case "open":
        return this.group();
      case "arrayOpen":
        return this.array();
      case "structured": {
        if (token.unterminated) {
          this.fail(token, "Unterminated structured reference");
        }
        this.index += 1;
        const bracket = token.text.indexOf("[");
        return bracket > 0
          ? { kind: "structured", table: token.text.slice(0, bracket), text: token.text, from: at, to: at }
          : { kind: "structured", text: token.text, from: at, to: at };
      }
      case "external":
        this.index += 1;
        return { kind: "external", text: token.text, from: at, to: at };
      default:
        return this.fail(token);
    }
  }

  private numberValue(token: Token): number {
    const value = Number(token.text);
    if (!Number.isFinite(value)) {
      this.fail(token, `${token.text} is not a number a cell can hold`);
    }
    return value;
  }

  private name(token: Token, at: number): FormulaNode {
    const prefix = sheetPrefixOf(token.text);
    if (prefix === "") {
      return { kind: "name", name: token.text, from: at, to: at };
    }
    const { sheet, sheetEnd } = parseSheetPrefix(prefix);
    if (sheetEnd !== undefined) {
      this.fail(token, "A name cannot span sheets");
    }
    return { kind: "name", name: token.text.slice(prefix.length), sheet, from: at, to: at };
  }

  private call(): FormulaNode {
    const from = this.index;
    const name = this.peek().text.toUpperCase();
    this.index += 1;
    const open = this.peek();
    if (open.kind !== "open") {
      this.fail(open);
    }
    this.enter(open);
    this.index += 1;
    const args: FormulaNode[] = [];
    if (this.peek().kind !== "close") {
      for (;;) {
        const next = this.peek();
        if (next.kind === "argSep" || next.kind === "close") {
          args.push({ kind: "missing", from: this.index, to: this.index - 1 });
        } else {
          args.push(this.expression());
        }
        const separator = this.peek();
        if (separator.kind === "argSep") {
          this.index += 1;
          continue;
        }
        if (separator.kind !== "close") {
          this.fail(separator);
        }
        break;
      }
    }
    const to = this.index;
    this.index += 1;
    this.leave();
    return { kind: "function", name, args, from, to };
  }

  private group(): FormulaNode {
    const from = this.index;
    this.enter(this.peek());
    this.index += 1;
    const inner = this.expression();
    const close = this.peek();
    if (close.kind !== "close") {
      this.fail(close, close.kind === "end" ? "Missing \")\"" : undefined);
    }
    const to = this.index;
    this.index += 1;
    this.leave();
    return { kind: "paren", inner, from, to };
  }

  private arrayElement(): ArrayElement {
    const token = this.peek();
    switch (token.kind) {
      case "number":
        this.index += 1;
        return this.numberValue(token);
      case "operator": {
        const next = this.tokens[this.index + 1]!;
        if ((token.text === "-" || token.text === "+") && next.kind === "number") {
          this.index += 2;
          const value = this.numberValue(next);
          return token.text === "-" ? -value : value;
        }
        return this.fail(token, "An array constant holds numbers, text, TRUE/FALSE and errors only");
      }
      case "string":
        if (token.unterminated) {
          this.fail(token, "Unterminated string");
        }
        this.index += 1;
        return token.text.slice(1, -1).replaceAll('""', '"');
      case "boolean":
        this.index += 1;
        return token.text.toUpperCase() === "TRUE";
      case "error": {
        const code = CODE_BY_UPPER.get(token.text.toUpperCase());
        if (code === undefined) {
          this.fail(token);
        }
        this.index += 1;
        return { error: code };
      }
      default:
        return this.fail(token, "An array constant holds numbers, text, TRUE/FALSE and errors only");
    }
  }

  private array(): FormulaNode {
    const from = this.index;
    this.index += 1;
    const rows: ArrayElement[][] = [[]];
    for (;;) {
      rows.at(-1)!.push(this.arrayElement());
      const separator = this.peek();
      if (separator.kind === "arrayColSep") {
        this.index += 1;
        continue;
      }
      const row = rows.at(-1)!;
      if (separator.kind !== "arrayRowSep" && separator.kind !== "arrayClose") {
        this.fail(separator);
      }
      if (rows.length > 1 && row.length !== rows[0]!.length) {
        this.fail(separator, "Every row of an array constant must have the same number of values");
      }
      this.index += 1;
      if (separator.kind === "arrayClose") {
        return { kind: "array", rows, from, to: this.index - 1 };
      }
      rows.push([]);
    }
  }
}

/** Parses tokens.ts's stream of a canonical formula. */
function parseTokens(tokens: readonly Token[]): FormulaNode {
  return new Parser(tokens).parse();
}

/** Parses canonical formula text (no leading "="). Throws
 * FormulaSyntaxError. */
export function parseFormula(text: string): FormulaNode {
  return parseTokens(tokenize(text));
}

// --- the AST cache ------------------------------------------------------------------

interface CacheEntry {
  node: FormulaNode;
  origin: CellRef;
  /** The token kinds the AST was parsed from (see the module header). */
  kinds: string;
}

/** relativeFormulaKey → AST; a Map's insertion order is the LRU order. */
const astCache = new Map<string, CacheEntry>();

function deepFreeze<T>(value: T): T {
  if (typeof value === "object" && value !== null && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) {
      deepFreeze(child);
    }
  }
  return value;
}

/** Parses `text` as the formula of `host`, sharing the AST of every earlier
 * formula with the same relative form (see the module header). Returns the
 * syntax error rather than throwing it. */
export function parseFormulaAt(text: string, host: CellRef): { node: FormulaNode; origin: CellRef } | FormulaSyntaxError {
  const tokens = tokenize(text);
  const key = relativeKeyOf(tokens, host);
  const kinds = tokens.map((token) => token.kind).join(" ");
  const hit = astCache.get(key);
  if (hit && hit.kinds === kinds) {
    astCache.delete(key);
    astCache.set(key, hit);
    return { node: hit.node, origin: hit.origin };
  }
  let node: FormulaNode;
  try {
    node = parseTokens(tokens);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return error;
    }
    throw error;
  }
  if (!hit) {
    const origin = { row: host.row, col: host.col };
    astCache.set(key, { node: deepFreeze(node), origin: deepFreeze(origin), kinds });
    if (astCache.size > FORMULA_AST_CACHE_SIZE) {
      astCache.delete(astCache.keys().next().value!);
    }
    return { node, origin };
  }
  return { node, origin: { row: host.row, col: host.col } };
}

/** How many ASTs the cache holds (for the specs and diagnostics). */
export function formulaAstCacheSize(): number {
  return astCache.size;
}

/** Empties the AST cache (for the specs). */
export function clearFormulaAstCache(): void {
  astCache.clear();
}
