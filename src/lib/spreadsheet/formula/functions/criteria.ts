// Criteria and wildcard matching (SPEC-120 "Criteria") — what the *IF/*IFS
// functions, the exact-match lookups (MATCH 0, VLOOKUP/HLOOKUP FALSE) and
// SEARCH match with.
//
// A criterion is a value. A number matches number cells equal at 15
// significant digits and text cells whose valueFromText is equal; a boolean
// matches that boolean; an empty-cell criterion is the number 0. A text
// criterion is an optional operator (=, <>, <, <=, >, >=) and an operand:
//
// | Operand is…               | `=` (or none), `<>`                  | `<` `<=` `>` `>=`            |
// | valueFromText succeeds    | numeric equality (+ numeric text)    | numeric order, number cells  |
// | TRUE / FALSE              | booleans                             | booleans                     |
// | an error literal          | errors of that code                  | nothing                      |
// | empty ("" or "=")         | "": empty cells and "" results;      | —                            |
// |                           | "=": empty cells; "<>": non-empty    |                              |
// | other text                | text cells, wildcards, case per      | collator order, text cells   |
// |                           | `caseSensitive`                      |                              |
// `<>` is the negation of `=` over every cell, empty ones included.
//
// Wildcards: `*` any run, `?` one UTF-16 unit, `~*` `~?` `~~` literals. The
// matcher is the greedy two-pointer algorithm with one backtrack point —
// O(text × pattern) at worst — and NO RegExp is ever built from cell or
// formula text. Case folding is per UTF-16 unit, so a match position is a
// position in the original text. The `.ods` settings adjust it: "plain"
// criteria match literally, "regex" criteria match literally when they hold
// none of `.*+?()[]{}|^$\` and are uncomputable otherwise, and a
// partial-cell-match workbook matches `=` text criteria as substrings.
import { ERROR_CODES, type ErrorCode } from "../../model";
import type { EvalSettings } from "../evaluate";
import { UncomputableError } from "../registry";
import { collatorFor, EMPTY, isEmpty, isError, roundTo15, valueFromText, type ScalarValue } from "../values";

export interface Criterion {
  matches(value: ScalarValue): boolean;
  /** Whether an empty cell matches — counted from an area's size. */
  readonly matchesEmpty: boolean;
}

const STAR = -1;
const ANY = -2;

/** A UTF-16 unit lower-cased when that keeps it one unit. */
function foldUnit(code: number): number {
  if (code < 0x41) {
    return code;
  }
  if (code <= 0x5a) {
    return code + 0x20;
  }
  if (code < 0x80) {
    return code;
  }
  const lower = String.fromCharCode(code).toLowerCase();
  return lower.length === 1 ? lower.charCodeAt(0) : code;
}

/** A text's units, folded unless case-sensitive. */
function units(text: string, caseSensitive: boolean): Int32Array {
  const out = new Int32Array(text.length);
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    out[i] = caseSensitive ? code : foldUnit(code);
  }
  return out;
}

/** A text folded unit by unit (see foldUnit). */
function foldText(text: string): string {
  let out = "";
  for (let i = 0; i < text.length; i += 1) {
    out += String.fromCharCode(foldUnit(text.charCodeAt(i)));
  }
  return out;
}

/** A pattern's `~*`, `~?` and `~~` escapes resolved to their literals. */
function unescapePattern(pattern: string): string {
  let out = "";
  for (let i = 0; i < pattern.length; i += 1) {
    const next = pattern[i + 1];
    if (pattern[i] === "~" && (next === "*" || next === "?" || next === "~")) {
      i += 1;
    }
    out += pattern[i];
  }
  return out;
}

/** A wildcard pattern as units: STAR, ANY or a literal (after `~`). */
function compilePattern(pattern: string, caseSensitive: boolean): Int32Array {
  const out: number[] = [];
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i]!;
    if (ch === "~" && i + 1 < pattern.length && (pattern[i + 1] === "*" || pattern[i + 1] === "?" || pattern[i + 1] === "~")) {
      i += 1;
      out.push(pattern.charCodeAt(i));
    } else if (ch === "*") {
      if (out.at(-1) !== STAR) {
        out.push(STAR);
      }
    } else if (ch === "?") {
      out.push(ANY);
    } else {
      const code = pattern.charCodeAt(i);
      out.push(caseSensitive ? code : foldUnit(code));
    }
  }
  return Int32Array.from(out);
}

/** The two-pointer match of text[from…] against a compiled pattern (the
 * whole remaining text). */
function matchFrom(text: Int32Array, from: number, pattern: Int32Array): boolean {
  let t = from;
  let p = 0;
  let starP = -1;
  let starT = 0;
  while (t < text.length) {
    const unit = pattern[p];
    if (p < pattern.length && unit !== STAR && (unit === ANY || unit === text[t])) {
      t += 1;
      p += 1;
    } else if (p < pattern.length && unit === STAR) {
      starP = p;
      p += 1;
      starT = t;
    } else if (starP >= 0) {
      p = starP + 1;
      starT += 1;
      t = starT;
    } else {
      return false;
    }
  }
  while (p < pattern.length && pattern[p] === STAR) {
    p += 1;
  }
  return p === pattern.length;
}

/** Whether a pattern holds an unescaped `*` or `?`. */
export function hasWildcards(pattern: string): boolean {
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i];
    if (ch === "~") {
      i += 1;
    } else if (ch === "*" || ch === "?") {
      return true;
    }
  }
  return false;
}

/** Whether the whole of `text` matches a wildcard pattern. */
export function wildcardMatch(text: string, pattern: string, caseSensitive = false): boolean {
  return matchFrom(units(text, caseSensitive), 0, compilePattern(pattern, caseSensitive));
}

const REGEX_SPECIALS = /[.*+?()[\]{}|^$\\]/;

/** How text is compared under the workbook's criteria setting: with
 * wildcards (xlsx; ods use-wildcards), else as the literal it spells — its
 * `~` escapes resolved under wildcards, verbatim for an `.ods` "plain" or
 * "regex" criterion. A regular expression is uncomputable. */
function textMode(pattern: string, settings: Pick<EvalSettings, "criteria">): { wildcards: true } | { wildcards: false; literal: string } {
  if (settings.criteria === "wildcards") {
    return hasWildcards(pattern) ? { wildcards: true } : { wildcards: false, literal: unescapePattern(pattern) };
  }
  if (settings.criteria === "regex" && REGEX_SPECIALS.test(pattern)) {
    throw new UncomputableError({ kind: "regex" });
  }
  return { wildcards: false, literal: pattern };
}

/** A text test: whole text (or, `partial`, a substring) equal to the
 * pattern under the settings' wildcard and case rules. */
export function textMatcher(
  pattern: string,
  settings: Pick<EvalSettings, "criteria" | "caseSensitive">,
  partial = false,
): (text: string) => boolean {
  const caseSensitive = settings.caseSensitive;
  const mode = textMode(pattern, settings);
  if (mode.wildcards) {
    const compiled = compilePattern(partial ? `*${pattern}*` : pattern, caseSensitive);
    return (text) => matchFrom(units(text, caseSensitive), 0, compiled);
  }
  const fold = (text: string): string => (caseSensitive ? text : foldText(text));
  const literal = fold(mode.literal);
  if (partial) {
    return (text) => fold(text).includes(literal);
  }
  return (text) => text.length === literal.length && fold(text) === literal;
}

/** The first 0-based position ≥ `start` where `find` matches a prefix of
 * the rest of `within` (SEARCH: case-insensitive, wildcards per the
 * settings); -1 when there is none. */
export function searchText(within: string, find: string, start: number, settings: Pick<EvalSettings, "criteria">): number {
  const mode = textMode(find, settings);
  if (!mode.wildcards) {
    return foldText(within).indexOf(foldText(mode.literal), start);
  }
  const text = units(within, false);
  const pattern = compilePattern(`${find}*`, false);
  for (let i = start; i <= text.length; i += 1) {
    if (matchFrom(text, i, pattern)) {
      return i;
    }
  }
  return -1;
}

type Operator = "=" | "<>" | "<" | "<=" | ">" | ">=";

/** A text criterion's leading operator (longest first) and its operand. */
function splitOperator(text: string): { op: Operator | null; operand: string } {
  for (const op of ["<=", ">=", "<>", "=", "<", ">"] as const) {
    if (text.startsWith(op)) {
      return { op, operand: text.slice(op.length) };
    }
  }
  return { op: null, operand: text };
}

function ordered(op: Operator, order: number): boolean {
  switch (op) {
    case "<":
      return order < 0;
    case "<=":
      return order <= 0;
    case ">":
      return order > 0;
    default:
      return order >= 0;
  }
}

/** `=` matches and `<>` negates over every cell; an order compares within
 * the type `same` accepts. */
function build(op: Operator, equals: (value: ScalarValue) => boolean, order?: (value: ScalarValue) => number | null): Criterion {
  let matches: (value: ScalarValue) => boolean;
  if (op === "=") {
    matches = equals;
  } else if (op === "<>") {
    matches = (value) => !equals(value);
  } else if (order) {
    matches = (value) => {
      const result = order(value);
      return result !== null && ordered(op, result);
    };
  } else {
    matches = () => false;
  }
  return { matches, matchesEmpty: matches(EMPTY) };
}

function numberCriterion(op: Operator, n: number, settings: EvalSettings): Criterion {
  const target = roundTo15(n);
  const asNumber = (value: ScalarValue): number | null => {
    if (typeof value === "number") {
      return roundTo15(value);
    }
    if (typeof value === "string") {
      const parsed = valueFromText(value, settings.locale, settings.dateSystem);
      return parsed === null ? null : roundTo15(parsed);
    }
    return null;
  };
  return build(
    op,
    (value) => asNumber(value) === target,
    (value) => (typeof value === "number" ? Math.sign(roundTo15(value) - target) : null),
  );
}

function booleanCriterion(op: Operator, b: boolean): Criterion {
  return build(
    op,
    (value) => value === b,
    (value) => (typeof value === "boolean" ? Number(value) - Number(b) : null),
  );
}

const ERROR_BY_UPPER: ReadonlyMap<string, ErrorCode> = new Map(ERROR_CODES.map((code) => [code.toUpperCase(), code]));

function errorCriterion(op: Operator, code: ErrorCode): Criterion {
  return build(op, (value) => isError(value) && value.error === code);
}

function textCriterion(op: Operator, operand: string, settings: EvalSettings): Criterion {
  if (op === "=" || op === "<>") {
    const test = textMatcher(operand, settings, !settings.wholeCell);
    return build(op, (value) => typeof value === "string" && test(value));
  }
  const collator = collatorFor(settings);
  return build(op, () => false, (value) => (typeof value === "string" ? collator.compare(value, operand) : null));
}

/** A criterion value as a matcher (see the module header). Throws
 * UncomputableError for an `.ods` regular expression. */
export function parseCriterion(criterion: ScalarValue, settings: EvalSettings): Criterion {
  if (isEmpty(criterion)) {
    return numberCriterion("=", 0, settings);
  }
  if (typeof criterion === "number") {
    return numberCriterion("=", criterion, settings);
  }
  if (typeof criterion === "boolean") {
    return booleanCriterion("=", criterion);
  }
  if (isError(criterion)) {
    return errorCriterion("=", criterion.error);
  }
  const { op, operand } = splitOperator(criterion);
  if (operand === "") {
    if (op === null) {
      return build("=", (value) => isEmpty(value) || value === "");
    }
    if (op === "=") {
      return build("=", (value) => isEmpty(value));
    }
    if (op === "<>") {
      return build("=", (value) => !isEmpty(value));
    }
    return build(op, () => false);
  }
  const effective = op ?? "=";
  const n = valueFromText(operand, settings.locale, settings.dateSystem);
  if (n !== null) {
    return numberCriterion(effective, n, settings);
  }
  const upper = operand.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return booleanCriterion(effective, upper === "TRUE");
  }
  const code = ERROR_BY_UPPER.get(upper);
  if (code !== undefined) {
    return errorCriterion(effective, code);
  }
  return textCriterion(effective, operand, settings);
}

/** The exact-match test of a lookup value (MATCH 0, VLOOKUP/HLOOKUP
 * FALSE): numbers equal at 15 digits, booleans equal, text whole-cell with
 * wildcards and case per the settings; an empty cell never matches. */
export function lookupMatcher(lookup: ScalarValue, settings: EvalSettings): (value: ScalarValue) => boolean {
  if (typeof lookup === "number") {
    const target = roundTo15(lookup);
    return (value) => typeof value === "number" && roundTo15(value) === target;
  }
  if (typeof lookup === "boolean") {
    return (value) => value === lookup;
  }
  if (typeof lookup === "string") {
    const test = textMatcher(lookup, settings);
    return (value) => typeof value === "string" && test(value);
  }
  return () => false;
}
