// The reference model and the token-level reference rewrite (SPEC-119
// "References") — THE primitive every later spec moves references with:
// SPEC-124's adjustFormula and SPEC-125's cut-move build on
// rewriteReferences, copy and fill on shiftFormula, sheet rename (SPEC-121)
// on renameSheetInFormula.
//
// - rewriteReferences tokenizes once, asks `rewrite` about each `ref`
//   token, replaces only the text of the tokens it rewrote (through
//   formatReference; `null` becomes the token's sheet prefix as written plus
//   `#REF!`), and joins: whitespace and every untouched token survive
//   byte-for-byte, and a formula nothing touched comes back as the identical
//   string. No rewrite here is print(parse()).
// - shiftFormula: a relative row part moves by dRow, a relative column part
//   by dCol (whole-column references have no row part, whole-row ones no
//   column part); any part leaving the grid turns the whole reference into
//   `#REF!` (`Sheet2!#REF!` keeps its prefix) — the text Excel produces on
//   copy and on shared-formula expansion. A range whose corners cross is
//   normalised, as Excel shows it (`$A$5:A11`, not `A11:$A$5`).
// - renameSheetInFormula: every reference (3-D endpoints included), sheet-
//   scoped name and `Sheet!#REF!` naming `from` — compared case-insensitively
//   — names `to` instead, re-quoted per the quoting rule (print.ts).
//   External references name another workbook's sheets and are left alone.
// - relativeFormulaKey: the token texts joined with trivia dropped, each
//   `ref` token as its R1C1 form relative to the host (`R[-1]C[2]` relative,
//   `R5C3` absolute, `C[0]:C[0]` a whole column), sheet prefixes as written —
//   two cells share a key exactly when one is a shifted copy of the other
//   (parse.ts keys its AST cache by it).
// - findArrayAnchor (formula/anchors.ts): the cell whose arrayRange holds a
//   position, memoised per `sheet.rows` array.
import type { CellRef } from "../address";
import { MAX_COLS, MAX_ROWS } from "../model";
import { formatReference, formatSheetPrefix } from "./print";
import { joinTokens, parseSheetPrefix, sheetPrefixOf, tokenize, type Token } from "./tokens";

export { findArrayAnchor } from "./anchors";

/** One corner of a reference, 0-based: `row` absent = whole column(s),
 * `col` absent = whole row(s). */
export interface RefPart {
  row?: number;
  col?: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export interface ReferenceNode {
  /** As written, unquoted and un-escaped. */
  sheet?: string;
  /** 3-D: the last sheet of `Sheet1:Sheet3!`. */
  sheetEnd?: string;
  start: RefPart;
  /** Absent for a single cell. */
  end?: RefPart;
}

/** A reference's replacement: a new reference, `null` for `#REF!`, or
 * `undefined` to keep the token verbatim. */
export type RefRewrite = (ref: ReferenceNode) => ReferenceNode | null | undefined;

/** Rewrites token texts through `visit` (which returns the new text, or
 * undefined to keep the token); the identical string when nothing changed. */
function rewriteTokens(text: string, visit: (token: Token) => string | undefined): string {
  const tokens = tokenize(text);
  let changed = false;
  for (const token of tokens) {
    const next = visit(token);
    if (next !== undefined && next !== token.text) {
      token.text = next;
      changed = true;
    }
  }
  return changed ? joinTokens(tokens) : text;
}

/** The token-level reference rewrite (see the module header). */
export function rewriteReferences(text: string, rewrite: RefRewrite): string {
  return rewriteTokens(text, (token) => {
    if (token.kind !== "ref" || !token.ref) {
      return undefined;
    }
    const next = rewrite(token.ref);
    if (next === undefined) {
      return undefined;
    }
    return next === null ? `${sheetPrefixOf(token.text)}#REF!` : formatReference(next);
  });
}

/** One part moved by (dRow, dCol): the same object when nothing moves, null
 * when it leaves the grid. */
function shiftPart(part: RefPart, dRow: number, dCol: number): RefPart | null {
  const row = part.row !== undefined && !part.rowAbs ? part.row + dRow : part.row;
  const col = part.col !== undefined && !part.colAbs ? part.col + dCol : part.col;
  if ((row !== undefined && (row < 0 || row >= MAX_ROWS)) || (col !== undefined && (col < 0 || col >= MAX_COLS))) {
    return null;
  }
  if (row === part.row && col === part.col) {
    return part;
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

/** A range's corners in order — the row parts swapped when the start row
 * passed the end row, the column parts likewise. */
function normalised(start: RefPart, end: RefPart): [RefPart, RefPart] {
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
  return [a, b];
}

/** A reference moved by (dRow, dCol) per shiftFormula: undefined when no
 * part moves, null when any part leaves the grid. */
export function shiftReference(ref: ReferenceNode, dRow: number, dCol: number): ReferenceNode | null | undefined {
  const start = shiftPart(ref.start, dRow, dCol);
  const end = ref.end ? shiftPart(ref.end, dRow, dCol) : undefined;
  if (start === null || end === null) {
    return null;
  }
  if (start === ref.start && end === ref.end) {
    return undefined;
  }
  const moved: ReferenceNode = { ...ref, start };
  if (end) {
    [moved.start, moved.end] = normalised(start, end);
  }
  return moved;
}

/** `text` with every relative reference moved by (dRow, dCol) (see the
 * module header). */
export function shiftFormula(text: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) {
    return text;
  }
  return rewriteReferences(text, (ref) => shiftReference(ref, dRow, dCol));
}

/** `text` with sheet `from` (any case) renamed `to` in every reference,
 * sheet-scoped name and `Sheet!#REF!` (see the module header). */
export function renameSheetInFormula(text: string, from: string, to: string): string {
  const key = from.toUpperCase();
  const renamed = (sheet: string | undefined): string | undefined =>
    sheet !== undefined && sheet.toUpperCase() === key ? to : sheet;
  return rewriteTokens(text, (token) => {
    if (token.kind === "ref" && token.ref) {
      const { sheet, sheetEnd } = token.ref;
      if (renamed(sheet) === sheet && renamed(sheetEnd) === sheetEnd) {
        return undefined;
      }
      const ref: ReferenceNode = { ...token.ref };
      ref.sheet = renamed(sheet);
      if (sheetEnd !== undefined) {
        ref.sheetEnd = renamed(sheetEnd);
      }
      return formatReference(ref);
    }
    if (token.kind !== "name" && token.kind !== "error" && token.kind !== "function") {
      return undefined;
    }
    const prefix = sheetPrefixOf(token.text);
    if (prefix === "") {
      return undefined;
    }
    const { sheet, sheetEnd } = parseSheetPrefix(prefix);
    if (renamed(sheet) === sheet && renamed(sheetEnd) === sheetEnd) {
      return undefined;
    }
    return formatSheetPrefix(renamed(sheet)!, renamed(sheetEnd)) + token.text.slice(prefix.length);
  });
}

/** One part in R1C1 form relative to `host`. */
function relativePart(part: RefPart, host: CellRef): string {
  let text = "";
  if (part.row !== undefined) {
    text += part.rowAbs ? `R${part.row + 1}` : `R[${part.row - host.row}]`;
  }
  if (part.col !== undefined) {
    text += part.colAbs ? `C${part.col + 1}` : `C[${part.col - host.col}]`;
  }
  return text;
}

/** relativeFormulaKey over an already-tokenized formula. */
export function relativeKeyOf(tokens: readonly Token[], host: CellRef): string {
  let key = "";
  for (const token of tokens) {
    if (token.kind === "ref" && token.ref) {
      const ref = token.ref;
      key += sheetPrefixOf(token.text) + relativePart(ref.start, host) + (ref.end ? `:${relativePart(ref.end, host)}` : "");
    } else {
      key += token.text;
    }
  }
  return key;
}

/** The formula's relative form at `host` (see the module header). */
export function relativeFormulaKey(text: string, host: CellRef): string {
  return relativeKeyOf(tokenize(text), host);
}
