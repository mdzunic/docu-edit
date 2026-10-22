// A conformance row (SPEC-127 "Conformance tables"): one formula, its inputs
// and the value Excel gives it, run by formula/conformance.spec.ts through
// the real parser, evaluator and recalculation. Rows live one file per
// category beside this one. Where an expected value may come from:
//
// - "Excel": Microsoft's documented behaviour — the function's reference
//   page and its example table (an Excel row wherever the page carries one),
//   `digits` set to the significant digits the page prints;
// - "pinned": computed by hand from the rule SPEC-120 or SPEC-127 states,
//   the arithmetic in `note`;
// - "reference": a published numerical table (Abramowitz & Stegun, NIST
//   DLMF) for the special functions.
//
// Never from running this engine, LibreOffice, Google Sheets or a JavaScript
// library: a copied bug is still a bug, and agreeing with ourselves proves
// nothing.
import type { CellValue } from "../../model";

export interface ConformanceRow {
  /** Canonical, no "=": "PMT(0.08/12,10,10000)". It is entered at Z100 of
   * Sheet1 (a legacy array formula at its range's top-left instead). */
  formula: string;
  /** Inputs on Sheet1 by A1 address; a text starting with "=" is a formula
   * (as calcWorkbook reads it), for FORMULATEXT and ISFORMULA. */
  cells?: Record<string, CellValue>;
  /** `{ error: "#NUM!" }` for an error row; the anchor's value for an array
   * row. */
  expected: CellValue;
  /** Entered as a legacy array formula over this range ("D1:E2"). */
  array?: string;
  /** Every member's value, row-major, when `array` is set. */
  expectedArray?: CellValue[][];
  /** Significant digits compared (default 15, as roundTo15). */
  digits?: number;
  source: "Excel" | "pinned" | "reference";
  /** Microsoft's page title, the derivation, or the table cited. */
  note: string;
}
