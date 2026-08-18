// SPEC-056 linear input parser: a LaTeX subset → MathNode[]. The linear form
// is an input method, never storage; the model is the MathNode tree. Input
// outside the table is inserted as literal `sym` nodes rather than raising an
// error, so a stray backslash never blocks typing.
import type { MathNode } from "../documentModel";
import {
  ACCENT_ALIASES,
  FUNCTION_NAMES,
  naryGlyphForName,
  resolveSymbolName,
  variantForText,
} from "./symbols";

interface Cursor {
  src: string;
  i: number;
}

const DIGITS = /[0-9]/;
const LETTER = /[A-Za-z]/;

/** Parse a LaTeX-subset string into a row of math nodes. Never throws. */
export function parseLinear(src: string): MathNode[] {
  const cur: Cursor = { src, i: 0 };
  return parseRow(cur, "");
}

function peek(cur: Cursor): string {
  return cur.i < cur.src.length ? cur.src[cur.i] : "";
}

function atStop(cur: Cursor, stops: string): boolean {
  if (cur.i >= cur.src.length) return true;
  const c = cur.src[cur.i];
  if (stops.includes("}") && c === "}") return true;
  if (stops.includes("&") && c === "&") return true;
  if (stops.includes("]") && c === "]") return true;
  if (stops.includes("|") && c === "|") return true;
  if (stops.includes("R") && startsWith(cur, "\\right")) return true;
  if (stops.includes("R") && startsWith(cur, "\\\\")) return true;
  if (stops.includes("&") && startsWith(cur, "\\\\")) return true;
  return false;
}

function startsWith(cur: Cursor, s: string): boolean {
  return cur.src.startsWith(s, cur.i);
}

function parseRow(cur: Cursor, stops: string): MathNode[] {
  const row: MathNode[] = [];
  while (!atStop(cur, stops)) {
    // Skip insignificant whitespace between atoms.
    if (peek(cur) === " ") {
      cur.i++;
      continue;
    }
    const unit = parsePrimary(cur);
    if (unit.length === 0) {
      // Nothing consumed (e.g. a lone closing token) — advance to avoid a
      // hang and drop the stray character as a literal upright sym.
      const ch = peek(cur);
      if (ch) {
        cur.i++;
        row.push({ kind: "sym", text: ch, variant: "upright" });
      }
      continue;
    }
    let sub: MathNode[] | undefined;
    let sup: MathNode[] | undefined;
    while (peek(cur) === "_" || peek(cur) === "^") {
      const which = cur.src[cur.i];
      cur.i++;
      const arg = parseScriptArg(cur);
      if (which === "_") sub = arg;
      else sup = arg;
    }
    if (sub || sup) {
      const single = unit.length === 1 ? unit[0] : null;
      if (single && single.kind === "nary") {
        if (sub) single.lower = sub;
        if (sup) single.upper = sup;
        row.push(single);
      } else {
        row.push({ kind: "script", base: unit, sup, sub });
      }
    } else {
      row.push(...unit);
    }
  }
  return row;
}

/** The argument of a `^` or `_`: a single group or a single primary token. */
function parseScriptArg(cur: Cursor): MathNode[] {
  if (peek(cur) === " ") cur.i++;
  if (peek(cur) === "{") {
    return parseGroup(cur);
  }
  return parseSingleToken(cur);
}

/** Parse a single primary token without following scripts. Returns the group
 * contents for `{…}`, or a one-element list for a command / character. */
function parseSingleToken(cur: Cursor): MathNode[] {
  if (peek(cur) === "{") return parseGroup(cur);
  if (peek(cur) === "\\") return parseCommand(cur);
  const ch = peek(cur);
  if (ch === "") return [];
  cur.i++;
  return [{ kind: "sym", text: ch, variant: variantForText(ch) }];
}

function parseGroup(cur: Cursor): MathNode[] {
  cur.i++; // consume "{"
  const nodes = parseRow(cur, "}");
  if (peek(cur) === "}") cur.i++;
  return nodes;
}

function parsePrimary(cur: Cursor): MathNode[] {
  const c = peek(cur);
  if (c === "{") return parseGroup(cur);
  if (c === "\\") return parseCommand(cur);
  if (DIGITS.test(c)) {
    let num = "";
    while (DIGITS.test(peek(cur))) {
      num += cur.src[cur.i];
      cur.i++;
    }
    return [{ kind: "sym", text: num, variant: "upright" }];
  }
  if (LETTER.test(c)) {
    cur.i++;
    return [{ kind: "sym", text: c, variant: "italic" }];
  }
  // A bracket that opens a plain delimiter pair (not via \left…\right).
  if (c === "[" || c === "(") {
    return parseBracketPair(cur, c);
  }
  // Any other literal character (operator, punctuation) is an upright sym.
  cur.i++;
  return [{ kind: "sym", text: c, variant: "upright" }];
}

function parseBracketPair(cur: Cursor, open: string): MathNode[] {
  const close = open === "[" ? "]" : ")";
  const stopChar = open === "[" ? "]" : "";
  cur.i++; // consume opener
  // For "(" there is no dedicated stop flag; scan until the matching ")".
  const body = open === "[" ? parseRow(cur, "]") : parseParenBody(cur);
  if (peek(cur) === close) cur.i++;
  else if (stopChar && peek(cur) === stopChar) cur.i++;
  return [{ kind: "delim", open, close, body }];
}

function parseParenBody(cur: Cursor): MathNode[] {
  const nodes: MathNode[] = [];
  while (cur.i < cur.src.length && peek(cur) !== ")") {
    if (peek(cur) === " ") {
      cur.i++;
      continue;
    }
    const unit = parsePrimary(cur);
    if (unit.length === 0) break;
    let sub: MathNode[] | undefined;
    let sup: MathNode[] | undefined;
    while (peek(cur) === "_" || peek(cur) === "^") {
      const which = cur.src[cur.i];
      cur.i++;
      const arg = parseScriptArg(cur);
      if (which === "_") sub = arg;
      else sup = arg;
    }
    if (sub || sup) {
      nodes.push({ kind: "script", base: unit, sup, sub });
    } else {
      nodes.push(...unit);
    }
  }
  return nodes;
}

function readCommandName(cur: Cursor): string {
  cur.i++; // consume backslash
  // A control symbol (e.g. \{ \| \\) is a single non-letter char.
  if (!LETTER.test(peek(cur))) {
    const ch = peek(cur);
    cur.i++;
    return ch;
  }
  let name = "";
  while (LETTER.test(peek(cur))) {
    name += cur.src[cur.i];
    cur.i++;
  }
  return name;
}

function parseCommand(cur: Cursor): MathNode[] {
  const start = cur.i;
  const name = readCommandName(cur);

  if (name === "frac" || name === "binom" || name === "dfrac" || name === "tfrac") {
    const num = parseArg(cur);
    const den = parseArg(cur);
    return [{ kind: "frac", num, den, bar: name !== "binom" }];
  }
  if (name === "sqrt") {
    let index: MathNode[] | undefined;
    if (peek(cur) === "[") {
      cur.i++;
      index = parseRow(cur, "]");
      if (peek(cur) === "]") cur.i++;
    }
    const radicand = parseArg(cur);
    return [{ kind: "radical", radicand, ...(index && index.length ? { index } : {}) }];
  }
  const nary = naryGlyphForName(name);
  if (nary) {
    return [{ kind: "nary", op: nary, body: [] }];
  }
  if (name === "left") {
    const open = readDelimiter(cur);
    const body = parseRow(cur, "R");
    let close = "";
    if (startsWith(cur, "\\right")) {
      cur.i += "\\right".length;
      close = readDelimiter(cur);
    }
    return [{ kind: "delim", open, close, body }];
  }
  if (name === "matrix" || name === "pmatrix" || name === "bmatrix") {
    const rows = parseMatrixBody(cur);
    if (name === "pmatrix") return [{ kind: "delim", open: "(", close: ")", body: [{ kind: "matrix", rows }] }];
    if (name === "bmatrix") return [{ kind: "delim", open: "[", close: "]", body: [{ kind: "matrix", rows }] }];
    return [{ kind: "matrix", rows }];
  }
  if (ACCENT_ALIASES[name]) {
    const base = parseArg(cur);
    return [{ kind: "accent", accent: ACCENT_ALIASES[name], base }];
  }
  if (name === "text" || name === "mathrm" || name === "operatorname") {
    const text = parseTextArg(cur);
    return [{ kind: "text", text }];
  }
  if (FUNCTION_NAMES.has(name)) {
    return [{ kind: "func", name, body: [] }];
  }
  const glyph = resolveSymbolName(name);
  if (glyph) {
    return [{ kind: "sym", text: glyph, variant: variantForText(glyph) }];
  }
  // Outside the table: literal sym nodes for the raw command text, so typing
  // never blocks (SPEC-056 "Input outside the linear input table").
  const raw = cur.src.slice(start, cur.i);
  return [...raw].map((ch) => ({ kind: "sym", text: ch, variant: "upright" as const }));
}

/** A required `{…}` argument; falls back to an empty slot when absent so an
 * incomplete `\frac{a}` still parses. */
function parseArg(cur: Cursor): MathNode[] {
  if (peek(cur) === " ") cur.i++;
  if (peek(cur) === "{") return parseGroup(cur);
  return parseSingleToken(cur);
}

/** `\text{…}` preserves its content verbatim (spaces included). */
function parseTextArg(cur: Cursor): string {
  if (peek(cur) !== "{") return "";
  cur.i++;
  let out = "";
  while (cur.i < cur.src.length && peek(cur) !== "}") {
    out += cur.src[cur.i];
    cur.i++;
  }
  if (peek(cur) === "}") cur.i++;
  return out;
}

function readDelimiter(cur: Cursor): string {
  if (peek(cur) === " ") cur.i++;
  if (peek(cur) === "\\") {
    const name = readCommandName(cur);
    if (name === "{") return "{";
    if (name === "}") return "}";
    if (name === "|") return "‖";
    if (name === "langle") return "⟨";
    if (name === "rangle") return "⟩";
    if (name === "lceil") return "⌈";
    if (name === "rceil") return "⌉";
    if (name === "lfloor") return "⌊";
    if (name === "rfloor") return "⌋";
    return name;
  }
  const ch = peek(cur);
  cur.i++;
  if (ch === ".") return ""; // \left. / \right. = no delimiter
  return ch;
}

function parseMatrixBody(cur: Cursor): MathNode[][][] {
  if (peek(cur) === " ") cur.i++;
  if (peek(cur) !== "{") return [[[]]];
  cur.i++;
  const rows: MathNode[][][] = [];
  let row: MathNode[][] = [];
  // Parse cells separated by "&" and rows by "\\", up to the closing "}".
  for (;;) {
    const cell = parseRow(cur, "}&");
    row.push(cell);
    if (peek(cur) === "&") {
      cur.i++;
      continue;
    }
    if (startsWith(cur, "\\\\")) {
      cur.i += 2;
      rows.push(row);
      row = [];
      continue;
    }
    break;
  }
  if (row.length) rows.push(row);
  if (peek(cur) === "}") cur.i++;
  return rows;
}
