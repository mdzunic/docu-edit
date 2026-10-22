// Point mode's text logic (SPEC-122 "Point mode") — THE normative source for
// where the cell editor may insert a pointed-at reference, how the
// references in a formula draft are found and coloured, and how F4 cycles a
// reference's `$` markers. Pure functions over SPEC-119's tokenize (the
// formula body after the leading "=" is tokenized; every offset here is in
// the whole editor text, "=" included):
//
// - canPointAt: the text starts with "=" and the last token before the
//   caret (trailing whitespace ignored) is "=" itself (nothing else yet),
//   "(", a "," (argument or union separator), a binary operator
//   (+ - * / ^ & = < > <= >= <>) or a unary sign. A "," inside an array
//   constant, ":", "%", a string, a reference or a name never allow it.
// - referenceSpans: every `ref` token of a formula draft, with its colour
//   index — the order of first appearance of each distinct reference text
//   (compared case-insensitively, `$` markers and sheet prefix included),
//   modulo the eight `--sheet-ref-1..8` colours. Non-formula text has none.
// - placePointReference: the pointed reference replaces the current point
//   span, or is inserted at the caret when there is none yet; the returned
//   span is the one the next arrow key or click replaces.
// - cycleAbsolute: the reference at or around the caret cycles A1 → $A$1 →
//   A$1 → $A1 → A1; a range's ends move together, taking the cycle position
//   from its first end (as Excel does); a whole column or row toggles the
//   only marker it shows. The sheet prefix stays as written; the caret lands
//   after the rewritten reference. Text without a reference there is
//   returned unchanged.
import { formatReference } from "./formula/print";
import type { ReferenceNode, RefPart } from "./formula/references";
import { tokenize, sheetPrefixOf, type Token } from "./formula/tokens";

/** How many reference colours the editor cycles through. */
export const REFERENCE_COLOR_COUNT = 8;

/** One reference in a formula draft: [start, end) in the editor text. */
export interface ReferenceSpan {
  start: number;
  end: number;
  ref: ReferenceNode;
  colorIndex: number;
}

const POINT_OPERATORS: ReadonlySet<string> = new Set(["+", "-", "*", "/", "^", "&", "=", "<", ">", "<=", ">=", "<>"]);

/** The formula body's tokens with offsets in the whole text ("=" at 0). */
function bodyTokens(text: string, end = text.length): Token[] {
  return tokenize(text.slice(1, end)).map((token) => ({ ...token, start: token.start + 1 }));
}

/** Whether a reference may be pointed at the caret (see the module header). */
export function canPointAt(text: string, caret: number): boolean {
  if (!text.startsWith("=") || caret < 1 || caret > text.length) {
    return false;
  }
  const tokens = tokenize(text.slice(1, caret));
  const last = tokens.length >= 2 ? tokens[tokens.length - 2]! : undefined;
  if (!last) {
    return true;
  }
  switch (last.kind) {
    case "open":
    case "argSep":
    case "union":
      return true;
    case "operator":
      return POINT_OPERATORS.has(last.text);
    default:
      return false;
  }
}

/** Every reference in a formula draft with its colour (see the header). */
export function referenceSpans(text: string): ReferenceSpan[] {
  if (!text.startsWith("=")) {
    return [];
  }
  const colors = new Map<string, number>();
  const spans: ReferenceSpan[] = [];
  for (const token of bodyTokens(text)) {
    if (token.kind !== "ref" || !token.ref) {
      continue;
    }
    const key = token.text.toUpperCase();
    let order = colors.get(key);
    if (order === undefined) {
      order = colors.size;
      colors.set(key, order);
    }
    spans.push({ start: token.start, end: token.start + token.text.length, ref: token.ref, colorIndex: order % REFERENCE_COLOR_COUNT });
  }
  return spans;
}

/** Inserts or replaces the pointed reference (see the module header): with
 * `span` null it goes in at `caret` (default: the end of the text). */
export function placePointReference(
  text: string,
  span: { start: number; end: number } | null,
  ref: string,
  caret: number = text.length,
): { text: string; span: { start: number; end: number } } {
  const start = Math.max(0, Math.min(span ? span.start : caret, text.length));
  const end = Math.max(start, Math.min(span ? span.end : start, text.length));
  return { text: text.slice(0, start) + ref + text.slice(end), span: { start, end: start + ref.length } };
}

/** The next `$` state: A1 → $A$1 → A$1 → $A1 → A1. */
function nextState(part: RefPart): { colAbs: boolean; rowAbs: boolean } {
  if (!part.colAbs && !part.rowAbs) {
    return { colAbs: true, rowAbs: true };
  }
  if (part.colAbs && part.rowAbs) {
    return { colAbs: false, rowAbs: true };
  }
  if (part.rowAbs) {
    return { colAbs: true, rowAbs: false };
  }
  return { colAbs: false, rowAbs: false };
}

/** The reference token at or around the caret: the one the caret is inside
 * or just after, else the one it is just before. */
function referenceAt(tokens: readonly Token[], caret: number): Token | undefined {
  let before: Token | undefined;
  for (const token of tokens) {
    if (token.kind !== "ref" || !token.ref) {
      continue;
    }
    const end = token.start + token.text.length;
    if (token.start < caret && caret <= end) {
      return token;
    }
    if (token.start === caret) {
      before = token;
    }
  }
  return before;
}

/** F4: cycles the `$` markers of the reference at the caret (see the
 * module header). */
export function cycleAbsolute(text: string, caret: number): { text: string; caret: number } {
  if (!text.startsWith("=")) {
    return { text, caret };
  }
  const token = referenceAt(bodyTokens(text), caret);
  if (!token?.ref) {
    return { text, caret };
  }
  const state = nextState(token.ref.start);
  const moved: ReferenceNode = { start: { ...token.ref.start, ...state } };
  if (token.ref.end) {
    moved.end = { ...token.ref.end, ...state };
  }
  const replacement = sheetPrefixOf(token.text) + formatReference(moved);
  const end = token.start + token.text.length;
  return { text: text.slice(0, token.start) + replacement + text.slice(end), caret: token.start + replacement.length };
}
