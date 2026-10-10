// Google Sheets' R1C1 formulas (SPEC-125 "External sources"): the
// `data-sheets-formula` a Sheets clipboard table carries, converted to
// canonical A1 text relative to each destination cell — `R[-1]C[0]` relative
// to the cell, `R2C3` absolute, `RC` the cell itself — string literals and
// quoted sheet names kept as written. A reference leaving the grid, a whole
// row or column (`R1:R3`, `C[1]`), or a result SPEC-119's parseFormula
// rejects makes the formula unconvertible: null, and the paste takes the
// cell's value instead.
import { colToLetters, type CellRef } from "../address";
import { MAX_COLS, MAX_ROWS } from "../model";
import { parseFormula } from "./parse";
import { printFormula } from "./print";

const R1C1_CELL = /^R(?:\[(-?\d+)\]|(\d+))?C(?:\[(-?\d+)\]|(\d+))?/i;
const R1C1_BAND = /^[RC](?:\[-?\d+\]|\d+)/i;
const IDENTIFIER = /[\p{L}\p{N}_.$]/u;

/** A Google Sheets R1C1 formula as canonical A1 text at `at` (SPEC-125
 * "External sources"): `R[-1]C[0]` relative to the cell, `R2C3` absolute;
 * null when a reference leaves the grid, names a whole row or column, or the
 * result does not parse. */
export function r1c1ToA1(formula: string, at: CellRef): string | null {
  const text = formula.startsWith("=") ? formula.slice(1) : formula;
  let out = "";
  let i = 0;
  while (i < text.length) {
    const ch = text[i]!;
    if (ch === '"' || ch === "'") {
      let end = i + 1;
      while (end < text.length && !(text[end] === ch && text[end + 1] !== ch)) {
        end += text[end] === ch ? 2 : 1;
      }
      out += text.slice(i, end + 1);
      i = end + 1;
      continue;
    }
    const boundary = i === 0 || !IDENTIFIER.test(text[i - 1]!);
    const match = boundary ? R1C1_CELL.exec(text.slice(i)) : null;
    const after = match ? text[i + match[0].length] : undefined;
    if (match && (after === undefined || (!IDENTIFIER.test(after) && after !== "("))) {
      const row = match[1] !== undefined ? at.row + Number(match[1]) : match[2] !== undefined ? Number(match[2]) - 1 : at.row;
      const col = match[3] !== undefined ? at.col + Number(match[3]) : match[4] !== undefined ? Number(match[4]) - 1 : at.col;
      if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) {
        return null;
      }
      out += `${match[4] !== undefined ? "$" : ""}${colToLetters(col)}${match[2] !== undefined ? "$" : ""}${row + 1}`;
      i += match[0].length;
      continue;
    }
    const band = boundary ? R1C1_BAND.exec(text.slice(i)) : null;
    const next = band ? text[i + band[0].length] : undefined;
    if (band && (next === undefined || (!IDENTIFIER.test(next) && next !== "("))) {
      return null;
    }
    out += ch;
    i += 1;
  }
  try {
    return printFormula(parseFormula(out));
  } catch {
    return null;
  }
}
