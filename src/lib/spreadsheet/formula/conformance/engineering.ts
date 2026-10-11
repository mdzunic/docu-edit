// Engineering conformance rows (SPEC-127 "Conformance tables"). Functions,
// in order: BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX, DEC2OCT, HEX2BIN,
// HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC, OCT2HEX, DELTA, GESTEP, ERF,
// ERF.PRECISE, ERFC, ERFC.PRECISE.
//
// Sources: each function's Microsoft reference page
// (support.microsoft.com/en-us/excel/functions/<name>-function) — its
// example table and Remarks bullets (the #NUM!/#VALUE! conditions are
// quoted there) — for "Excel" rows; SPEC-127's Function pins (10-character
// two's complement; DEC2BIN -512..511, DEC2OCT -2^29..2^29-1, DEC2HEX
// -2^39..2^39-1) for hand-derived "pinned" rows; Abramowitz & Stegun Table
// 7.1 (the error function) for "reference" rows. Never this engine's own
// output.
import type { ConformanceRow } from "./row";

export const ENGINEERING_ROWS: ConformanceRow[] = [
  // --- BIN2DEC: number ---------------------------------------------------
  { formula: "BIN2DEC(1100100)", expected: 100, source: "Excel", note: "BIN2DEC function, example, row 1." },
  { formula: "BIN2DEC(1111111111)", expected: -1, source: "Excel", note: "BIN2DEC function, example, row 2 (two's-complement, sign bit set)." },
  {
    formula: "BIN2DEC(11111111111)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "BIN2DEC Remarks, verbatim: \"If number ... contains more than 10 characters (10 bits), BIN2DEC returns the #NUM! error value.\" (11 digits here).",
  },

  // --- BIN2HEX: number, [places] ------------------------------------------
  { formula: "BIN2HEX(11111011,4)", expected: "00FB", source: "Excel", note: "BIN2HEX function, example, row 1." },
  { formula: "BIN2HEX(1110)", expected: "E", source: "Excel", note: "BIN2HEX function, example, row 2." },
  { formula: "BIN2HEX(1111111111)", expected: "FFFFFFFFFF", source: "Excel", note: "BIN2HEX function, example, row 3 (negative, 10-char result)." },
  {
    formula: "BIN2HEX(11111011,1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "BIN2HEX Remarks, verbatim: \"If BIN2HEX requires more than places characters, it returns the #NUM! error value.\" (needs 2 chars, \"FB\").",
  },

  // --- BIN2OCT: number, [places] ------------------------------------------
  { formula: "BIN2OCT(1001,3)", expected: "011", source: "Excel", note: "BIN2OCT function, example, row 1." },
  { formula: "BIN2OCT(1100100)", expected: "144", source: "Excel", note: "BIN2OCT function, example, row 2." },
  { formula: "BIN2OCT(1111111111)", expected: "7777777777", source: "Excel", note: "BIN2OCT function, example, row 3 (negative, 10-char result)." },
  {
    formula: "BIN2OCT(1100100,1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "BIN2OCT Remarks, verbatim: \"If BIN2OCT requires more than places characters, it returns the #NUM! error value.\" (needs 3 chars, \"144\").",
  },

  // --- DEC2BIN: number, [places] ------------------------------------------
  { formula: "DEC2BIN(9,4)", expected: "1001", source: "Excel", note: "DEC2BIN function, example, row 1." },
  { formula: "DEC2BIN(-100)", expected: "1110011100", source: "Excel", note: "DEC2BIN function, example, row 2." },
  {
    formula: "DEC2BIN(512)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "DEC2BIN Remarks, verbatim: \"If number < -512 or if number > 511, DEC2BIN returns the #NUM! error value.\"",
  },

  // --- DEC2HEX: number, [places] ------------------------------------------
  { formula: "DEC2HEX(100,4)", expected: "0064", source: "Excel", note: "DEC2HEX function, example, row 1." },
  { formula: "DEC2HEX(-54)", expected: "FFFFFFFFCA", source: "Excel", note: "DEC2HEX function, example, row 2." },
  { formula: "DEC2HEX(28)", expected: "1C", source: "Excel", note: "DEC2HEX function, example, row 3." },
  {
    formula: "DEC2HEX(64,1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "DEC2HEX function, example, row 4: Microsoft states the result (40 hex) needs 2 places, so places=1 is #NUM!.",
  },

  // --- DEC2OCT: number, [places] ------------------------------------------
  { formula: "DEC2OCT(58,3)", expected: "072", source: "Excel", note: "DEC2OCT function, example, row 1." },
  { formula: "DEC2OCT(-100)", expected: "7777777634", source: "Excel", note: "DEC2OCT function, example, row 2." },
  {
    formula: "DEC2OCT(536870912)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "DEC2OCT Remarks, verbatim: \"If number < -536,870,912 or if number > 536,870,911, DEC2OCT returns the #NUM! error value.\"",
  },

  // --- HEX2BIN: number, [places] ------------------------------------------
  { formula: 'HEX2BIN("F",8)', expected: "00001111", source: "Excel", note: "HEX2BIN function, example, row 1." },
  { formula: 'HEX2BIN("B7")', expected: "10110111", source: "Excel", note: "HEX2BIN function, example, row 2." },
  { formula: 'HEX2BIN("FFFFFFFFFF")', expected: "1111111111", source: "Excel", note: "HEX2BIN function, example, row 3." },
  {
    formula: 'HEX2BIN("G")',
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "HEX2BIN Remarks, verbatim: \"If number is not a valid hexadecimal number, HEX2BIN returns the #NUM! error value.\" (\"G\" is not a hex digit).",
  },

  // --- HEX2DEC: number -----------------------------------------------------
  { formula: 'HEX2DEC("A5")', expected: 165, source: "Excel", note: "HEX2DEC function, example, row 1." },
  { formula: 'HEX2DEC("FFFFFFFF5B")', expected: -165, source: "Excel", note: "HEX2DEC function, example, row 2 (two's complement)." },
  {
    formula: 'HEX2DEC("3DA408B9")',
    expected: 1034000000,
    digits: 4,
    source: "Excel",
    note: "HEX2DEC function, example, row 3; Microsoft prints 1.034E+09, i.e. 1034000000 to 4 significant digits.",
  },
  {
    formula: 'HEX2DEC("G")',
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "HEX2DEC Remarks, verbatim: \"If number is not a valid hexadecimal number, HEX2DEC returns the #NUM! error value.\"",
  },

  // --- HEX2OCT: number, [places] ------------------------------------------
  { formula: 'HEX2OCT("F",3)', expected: "017", source: "Excel", note: "HEX2OCT function, example, row 1." },
  { formula: 'HEX2OCT("3B4E")', expected: "35516", source: "Excel", note: "HEX2OCT function, example, row 2." },
  { formula: 'HEX2OCT("FFFFFFFF00")', expected: "7777777400", source: "Excel", note: "HEX2OCT function, example, row 3." },
  {
    formula: 'HEX2OCT("G")',
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "HEX2OCT Remarks, verbatim: \"If number is not a valid hexadecimal number, HEX2OCT returns the #NUM! error value.\"",
  },

  // --- OCT2BIN: number, [places] ------------------------------------------
  { formula: "OCT2BIN(3,3)", expected: "011", source: "Excel", note: "OCT2BIN function, example, row 1." },
  { formula: "OCT2BIN(7777777000)", expected: "1000000000", source: "Excel", note: "OCT2BIN function, example, row 2 (negative, 10-char result)." },
  {
    formula: "OCT2BIN(1000)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "OCT2BIN Remarks, verbatim: \"If number is positive, it cannot be greater than 777.\" (1000 > 777).",
  },

  // --- OCT2DEC: number -------------------------------------------------------
  { formula: "OCT2DEC(54)", expected: 44, source: "Excel", note: "OCT2DEC function, example, row 1." },
  { formula: "OCT2DEC(7777777533)", expected: -165, source: "Excel", note: "OCT2DEC function, example, row 2 (two's complement)." },
  {
    formula: "OCT2DEC(8)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "OCT2DEC Remarks, verbatim: \"If number is not a valid octal number, OCT2DEC returns the #NUM! error value.\" (8 is not an octal digit).",
  },

  // --- OCT2HEX: number, [places] ------------------------------------------
  { formula: "OCT2HEX(100,4)", expected: "0040", source: "Excel", note: "OCT2HEX function, example, row 1." },
  { formula: "OCT2HEX(7777777533)", expected: "FFFFFFFF5B", source: "Excel", note: "OCT2HEX function, example, row 2 (negative, 10-char result)." },
  {
    formula: "OCT2HEX(8)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "OCT2HEX Remarks, verbatim: \"If number is not a valid octal number, OCT2HEX returns the #NUM! error value.\" (8 is not an octal digit).",
  },

  // --- DELTA: number1, [number2] ------------------------------------------
  { formula: "DELTA(5,4)", expected: 0, source: "Excel", note: "DELTA function, example, row 1." },
  { formula: "DELTA(5,5)", expected: 1, source: "Excel", note: "DELTA function, example, row 2." },
  { formula: "DELTA(0.5,0)", expected: 0, source: "Excel", note: "DELTA function, example, row 3." },
  {
    formula: 'DELTA("a",1)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "DELTA Remarks, verbatim: \"If number1 is nonnumeric, DELTA returns the #VALUE! error value.\"",
  },

  // --- GESTEP: number, [step] ----------------------------------------------
  { formula: "GESTEP(5,4)", expected: 1, source: "Excel", note: "GESTEP function, example, row 1." },
  { formula: "GESTEP(5,5)", expected: 1, source: "Excel", note: "GESTEP function, example, row 2." },
  { formula: "GESTEP(-4,-5)", expected: 1, source: "Excel", note: "GESTEP function, example, row 3." },
  { formula: "GESTEP(-1)", expected: 0, source: "Excel", note: "GESTEP function, example, row 4: step omitted, defaults to 0." },
  {
    formula: 'GESTEP("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "GESTEP Remarks, verbatim: \"If any argument is nonnumeric, GESTEP returns the #VALUE! error value.\"",
  },

  // --- ERF: lower_limit, [upper_limit] -------------------------------------
  { formula: "ERF(0.745)", expected: 0.70792892, digits: 8, source: "Excel", note: "ERF function, example, row 1." },
  { formula: "ERF(1)", expected: 0.84270079, digits: 8, source: "Excel", note: "ERF function, example, row 2." },
  {
    formula: "ERF(0)",
    expected: 0,
    source: "reference",
    note: "reference: Abramowitz & Stegun Table 7.1, x=0.00: erf(0)=0 (erf is an odd function, trivially zero at the origin).",
  },
  {
    formula: 'ERF("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "ERF Remarks, verbatim: \"If lower_limit is nonnumeric, ERF returns the #VALUE! error value.\"",
  },

  // --- ERF.PRECISE: x -------------------------------------------------------
  { formula: "ERF.PRECISE(0.745)", expected: 0.70792892, digits: 8, source: "Excel", note: "ERF.PRECISE function, example, row 1 (same values as ERF)." },
  { formula: "ERF.PRECISE(1)", expected: 0.84270079, digits: 8, source: "Excel", note: "ERF.PRECISE function, example, row 2." },
  {
    formula: "ERF.PRECISE(0)",
    expected: 0,
    source: "reference",
    note: "reference: Abramowitz & Stegun Table 7.1, x=0.00: erf(0)=0.",
  },
  {
    formula: 'ERF.PRECISE("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "ERF.PRECISE Remarks, verbatim: \"If lower_limit is nonnumeric, ERF.PRECISE returns the #VALUE! error value.\"",
  },

  // --- ERFC: x ---------------------------------------------------------------
  { formula: "ERFC(1)", expected: 0.15729921, digits: 8, source: "Excel", note: "ERFC function, example: complementary ERF of 1." },
  {
    formula: "ERFC(0)",
    expected: 1,
    source: "reference",
    note: "reference: Abramowitz & Stegun Table 7.1, x=0.00: erfc(x)=1-erf(x)=1-0=1.",
  },
  {
    formula: "ERFC(0.745)",
    expected: 0.29207108,
    digits: 8,
    source: "pinned",
    note: "pinned: erfc(x)=1-erf(x); ERF's page gives erf(0.745)=0.70792892, so erfc(0.745)=1-0.70792892=0.29207108.",
  },
  {
    formula: 'ERFC("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "ERFC Remarks, verbatim: \"If x is nonnumeric, ERFC returns the #VALUE! error value.\"",
  },

  // --- ERFC.PRECISE: x ---------------------------------------------------------
  { formula: "ERFC.PRECISE(1)", expected: 0.15729921, digits: 8, source: "Excel", note: "ERFC.PRECISE function, example: complementary ERF of 1." },
  {
    formula: "ERFC.PRECISE(0)",
    expected: 1,
    source: "reference",
    note: "reference: Abramowitz & Stegun Table 7.1, x=0.00: erfc(x)=1-erf(x)=1-0=1.",
  },
  {
    formula: "ERFC.PRECISE(0.745)",
    expected: 0.29207108,
    digits: 8,
    source: "pinned",
    note: "pinned: erfc(x)=1-erf(x); ERF's page gives erf(0.745)=0.70792892, so erfc(0.745)=1-0.70792892=0.29207108.",
  },
  {
    formula: 'ERFC.PRECISE("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "ERFC.PRECISE Remarks, verbatim: \"If x is nonnumeric, ERFC.PRECISE returns the #VALUE! error value.\"",
  },
];
