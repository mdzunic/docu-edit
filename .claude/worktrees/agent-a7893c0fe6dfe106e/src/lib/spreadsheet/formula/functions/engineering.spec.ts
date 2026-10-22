// SPEC-127 Engineering — functions/engineering.ts: the twelve base
// conversions, DELTA and GESTEP, each evaluated in a cell of an in-memory
// workbook. "Excel" rows are the examples on Microsoft's reference pages;
// the others are pinned from the rules the module header states (10-character
// two's complement, places truncated, too few or negative places #NUM!,
// places ignored for a negative result, more than 10 characters or a stray
// digit #NUM!, upper-case output).
import { describe, expect, it } from "vitest";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext } from "../evaluate";

/** `formula` evaluated in Sheet1!Z100 over `cells` — the value it stores. */
function value(formula: string, cells: Record<string, CellSpec> = {}) {
  const model = calcWorkbook({ Sheet1: { ...cells, Z100: `=${formula}` } });
  return evaluateCell(model, "sheet-Sheet1", 99, 25, storedValueContext(model));
}

const err = (code: ErrorCode): CellValue => ({ error: code });

describe("the base conversions", () => {
  it.each<[string, CellValue]>([
    ["BIN2DEC(1100100)", 100],
    ["BIN2DEC(1111111111)", -1],
    ["BIN2HEX(11111011,4)", "00FB"],
    ["BIN2HEX(1110)", "E"],
    ["BIN2HEX(1111111111)", "FFFFFFFFFF"],
    ["BIN2OCT(1001,3)", "011"],
    ["BIN2OCT(1100100)", "144"],
    ["BIN2OCT(1111111111)", "7777777777"],
    ["DEC2BIN(9,4)", "1001"],
    ["DEC2BIN(-100)", "1110011100"],
    ["DEC2HEX(100,4)", "0064"],
    ["DEC2HEX(-54)", "FFFFFFFFCA"],
    ["DEC2HEX(28)", "1C"],
    ["DEC2HEX(64,1)", err("#NUM!")],
    ["DEC2OCT(58,3)", "072"],
    ["DEC2OCT(-100)", "7777777634"],
    ['HEX2BIN("F",8)', "00001111"],
    ['HEX2BIN("B7")', "10110111"],
    ['HEX2BIN("FFFFFFFFFF")', "1111111111"],
    ['HEX2DEC("A5")', 165],
    ['HEX2DEC("FFFFFFFF5B")', -165],
    ['HEX2DEC("3DA408B9")', 1034160313], // printed 1.034E+09; 0x3DA408B9 by hand
    ['HEX2OCT("F",3)', "017"],
    ['HEX2OCT("3B4E")', "35516"],
    ['HEX2OCT("FFFFFFFF00")', "7777777400"],
    ["OCT2BIN(3,3)", "011"],
    ["OCT2BIN(7777777000)", "1000000000"],
    ["OCT2DEC(54)", 44],
    ["OCT2DEC(7777777533)", -165],
    ["OCT2HEX(100,4)", "0040"],
    ["OCT2HEX(7777777533)", "FFFFFFFF5B"],
  ])("Excel: %s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it.each<[string, CellValue]>([
    // Ranges: DEC2BIN −512…511, DEC2OCT −2^29…2^29−1, DEC2HEX −2^39…2^39−1.
    ["DEC2BIN(511)", "111111111"],
    ["DEC2BIN(-512)", "1000000000"],
    ["DEC2BIN(512)", err("#NUM!")],
    ["DEC2BIN(-513)", err("#NUM!")],
    ["DEC2OCT(536870911)", "3777777777"],
    ["DEC2OCT(-536870912)", "4000000000"],
    ["DEC2OCT(536870912)", err("#NUM!")],
    ["DEC2HEX(549755813887)", "7FFFFFFFFF"],
    ["DEC2HEX(-549755813888)", "8000000000"],
    ["DEC2HEX(549755813888)", err("#NUM!")],
    ['HEX2BIN("1FF")', "111111111"],
    ['HEX2BIN("200")', err("#NUM!")],
    ['HEX2BIN("FFFFFFFE00")', "1000000000"],
    ['HEX2BIN("FFFFFFFDFF")', err("#NUM!")],
    ['HEX2OCT("FFE0000000")', "4000000000"],
    ['HEX2OCT("20000000")', err("#NUM!")],
    ["OCT2BIN(1000)", err("#NUM!")],
    // Places: truncated; too few, zero or negative #NUM!; ignored for a negative result.
    ["DEC2BIN(9,4.9)", "1001"],
    ["DEC2BIN(9,3)", err("#NUM!")],
    ["DEC2BIN(9,0)", err("#NUM!")],
    ["DEC2HEX(100,-1)", err("#NUM!")],
    ["DEC2BIN(-1,4)", "1111111111"],
    ['DEC2BIN(9,"x")', err("#VALUE!")],
    // Input digits: more than 10 characters, a stray digit, a sign or a point #NUM!.
    ['BIN2DEC("00000000001")', err("#NUM!")],
    ["BIN2DEC(102)", err("#NUM!")],
    ["BIN2DEC(-1)", err("#NUM!")],
    ["BIN2DEC(10.1)", err("#NUM!")],
    ['HEX2DEC("G1")', err("#NUM!")],
    ['HEX2DEC("a5")', 165], // lower case accepted
    ['HEX2BIN("a",4)', "1010"],
    ["OCT2DEC(8)", err("#NUM!")],
    ['BIN2DEC("")', 0],
    ["BIN2DEC(TRUE)", err("#VALUE!")],
    // A decimal number is coerced and truncated toward zero.
    ['DEC2BIN("9")', "1001"],
    ["DEC2BIN(9.9)", "1001"],
    ['DEC2HEX("a")', err("#VALUE!")],
    ["DEC2HEX(255)", "FF"], // upper case
  ])("pinned: %s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("reads a reference's number and text, and an empty cell as 0", () => {
    expect(value("BIN2DEC(A1)", { A1: 101 })).toBe(5);
    expect(value("HEX2DEC(A1)", { A1: "ff" })).toBe(255);
    expect(value("OCT2DEC(A2)", { A1: 1 })).toBe(0);
  });

  it("propagates an argument's error", () => {
    expect(value("DEC2BIN(1/0)")).toEqual(err("#DIV/0!"));
    expect(value("BIN2HEX(1,NA())")).toEqual(err("#N/A"));
  });
});

describe("DELTA and GESTEP", () => {
  it.each<[string, CellValue]>([
    ["DELTA(5,4)", 0],
    ["DELTA(5,5)", 1],
    ["DELTA(0.5,0)", 0],
    ["GESTEP(5,4)", 1],
    ["GESTEP(5,5)", 1],
    ["GESTEP(-4,-5)", 1],
    ["GESTEP(-1)", 0],
  ])("Excel: %s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it.each<[string, CellValue]>([
    ["DELTA(0)", 1], // number2 defaults to 0
    ["DELTA(0.1+0.2,0.3)", 1], // compared at 15 significant digits, as "=" does
    ['DELTA("a",1)', err("#VALUE!")],
    ["GESTEP(0)", 1],
    ['GESTEP(1,"b")', err("#VALUE!")],
    ["GESTEP(NA())", err("#N/A")],
  ])("pinned: %s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});
