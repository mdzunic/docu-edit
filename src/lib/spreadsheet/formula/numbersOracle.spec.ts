// SPEC-120 "The Numbers oracle": e2e/fixtures/spec120-numbers.xlsx is a
// third-party engine's arithmetic — Apple Numbers 14.4, run as en-US,
// exported to Excel (e2e/fixtures/spec120-numbers.md records it). The
// workbook is recalculated in full exactly as the product recalculates one,
// and every formula cell must match Numbers' cached value:
// - a formula cell with no <v> (fourteen of them) is one Numbers evaluated
//   to an error — it matches any CellError;
// - a t="b" result compares as a boolean, t="str" as exact text, a number
//   under roundTo15;
// - the nine rows of the note's "Where Numbers differs from Excel" table are
//   the listed exceptions, each asserting the Excel / SPEC-120 value the
//   table gives, with its reason.
// All 147 formula cells are single-cell dynamic-array formulas; they compute
// through "Array mode"'s single-cell rule, so none may be left uncomputable.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vitest";
import { formatA1 } from "../address";
import type { CellValue, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { recalculate, type RecalcResult } from "./recalc";
import { roundTo15 } from "./values";

const FIXTURE = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "..", "e2e", "fixtures", "spec120-numbers.xlsx");

/** The note's "Where Numbers differs from Excel" rows: the value Excel (and
 * this engine) gives, and why Numbers' differs. */
const EXCEPTIONS: Readonly<Record<string, { formula: string; expected: CellValue; reason: string }>> = {
  E35: { formula: '"z"<TRUE', expected: true, reason: "Numbers does not compare text with a boolean" },
  E51: { formula: "CEILING(-2.5,2)", expected: -2, reason: "Numbers' CEILING needs both arguments to have the same sign (Excel 2007's rule)" },
  E79: { formula: 'SUM("3",TRUE)', expected: 4, reason: "Numbers' SUM rejects a boolean argument" },
  E110: { formula: 'PROPER("it\'s")', expected: "It'S", reason: "Numbers does not capitalise after an apostrophe" },
  E121: { formula: 'LEN("😀")', expected: 2, reason: "Numbers counts characters, not UTF-16 code units" },
  E124: { formula: "DATE(108,1,2)", expected: 39449, reason: "Numbers takes year 108 literally (no +1900) and exports the displayed text" },
  E125: { formula: "DATE(1900,2,29)", expected: 60, reason: "Numbers' calendar has no 29 Feb 1900; Excel keeps the 1900 leap-year bug" },
  E129: { formula: "TIME(12,30,0)", expected: 0.520833333333333, reason: "Numbers anchors a time of day to 1 Jan 1904 (serial 1462)" },
  E135: { formula: "WEEKDAY(DATE(1900,1,1))", expected: 1, reason: "1 Jan 1900 was a Monday; Excel's leap-bug arithmetic makes it a Sunday" },
};

/** The fourteen cells Numbers left without a cached value (errors). */
const NUMBERS_ERRORS = ["E16", "E17", "E27", "E29", "E30", "E31", "E35", "E39", "E40", "E51", "E79", "E81", "E83", "E115"];

interface FormulaCell {
  address: string;
  formula: string;
  cached: CellValue;
}

function formulaCells(model: SpreadsheetModel): FormulaCell[] {
  const sheet = model.sheets[0]!;
  return sheet.rows.flatMap((row) =>
    row.cells.flatMap((cell) => (cell.formula === undefined ? [] : [{ address: formatA1({ row: row.index, col: cell.col }), formula: cell.formula, cached: cell.value }])),
  );
}

function valueAt(model: SpreadsheetModel, address: string): CellValue {
  for (const row of model.sheets[0]!.rows) {
    for (const cell of row.cells) {
      if (formatA1({ row: row.index, col: cell.col }) === address) {
        return cell.value;
      }
    }
  }
  return null;
}

/** Numbers' cached value against the engine's, per "The Numbers oracle". */
function matches(cached: CellValue, computed: CellValue): boolean {
  if (cached === null) {
    return typeof computed === "object" && computed !== null && "error" in computed;
  }
  if (typeof cached === "number") {
    return typeof computed === "number" && roundTo15(computed) === roundTo15(cached);
  }
  return computed === cached;
}

describe("the Numbers oracle (spec120-numbers.xlsx, Numbers 14.4 as en-US)", () => {
  let model: SpreadsheetModel;
  let before: FormulaCell[];
  let result: RecalcResult;

  beforeAll(async () => {
    model = await readSpreadsheetBytes(new Uint8Array(readFileSync(FIXTURE)), "xlsx", "spec120-numbers");
    before = formulaCells(model);
    result = recalculate(model, [], { full: true, reason: "manual", now: () => new Date(2026, 9, 9, 12, 0) });
  });

  it("holds 147 formula cells in column E, every one a single-cell dynamic array, and fourteen without a cached value", () => {
    expect(before).toHaveLength(147);
    expect(before.every((cell) => cell.address.startsWith("E"))).toBe(true);
    const sheet = model.sheets[0]!;
    for (const row of sheet.rows) {
      for (const cell of row.cells.filter((candidate) => candidate.formula !== undefined)) {
        expect(cell.dynamicArray, formatA1({ row: row.index, col: cell.col })).toBe(true);
        expect(cell.arrayRange).toEqual({ top: row.index, left: cell.col, bottom: row.index, right: cell.col });
      }
    }
    expect(before.filter((cell) => cell.cached === null).map((cell) => cell.address)).toEqual(NUMBERS_ERRORS);
  });

  it("computes every formula cell — none is left uncomputable, so none is stale", () => {
    expect(result.evaluated).toBe(147);
    expect(result.stale).toEqual([]);
    expect(result.circular).toEqual([]);
  });

  it("names the listed exceptions by their formulas, as the note's table does", () => {
    for (const [address, exception] of Object.entries(EXCEPTIONS)) {
      expect(before.find((cell) => cell.address === address)?.formula, address).toBe(exception.formula);
    }
  });

  it("matches Numbers' cached value in every other formula cell (no <v> → any error; booleans, exact text, numbers under roundTo15)", () => {
    const mismatches: string[] = [];
    for (const cell of before) {
      if (cell.address in EXCEPTIONS) {
        continue;
      }
      const computed = valueAt(model, cell.address);
      if (!matches(cell.cached, computed)) {
        mismatches.push(`${cell.address} =${cell.formula}: Numbers ${JSON.stringify(cell.cached)}, computed ${JSON.stringify(computed)}`);
      }
    }
    expect(mismatches).toEqual([]);
  });

  for (const [address, exception] of Object.entries(EXCEPTIONS)) {
    it(`${address} =${exception.formula} is the Excel / SPEC-120 value ${JSON.stringify(exception.expected)} — ${exception.reason}`, () => {
      const computed = valueAt(model, address);
      if (typeof exception.expected === "number") {
        expect(typeof computed === "number" ? roundTo15(computed) : computed).toBe(exception.expected);
      } else {
        expect(computed).toBe(exception.expected);
      }
    });
  }
});
