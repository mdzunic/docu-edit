// The Math & Trig conformance rows (SPEC-127 "Conformance tables").
// Sources: Microsoft's function reference pages (support.microsoft.com/en-us/excel/functions/<name>-function),
// hand derivation from SPEC-120/127's stated rules (noted "pinned"), and this
// file's own small hand-built datasets for SUBTOTAL/AGGREGATE/MMULT, whose
// worked examples on Microsoft's pages are not reliably machine-transcribable
// (an embedded image for MMULT; an inconsistently rendered data table for
// AGGREGATE) — see SPEC-127's brief, "SUBTOTAL and AGGREGATE": pinned rows on
// visible data are sanctioned there.
import type { ConformanceRow } from "./row";

export const MATH_ROWS: ConformanceRow[] = [
  // --- SUM ---------------------------------------------------------------
  { formula: "SUM(1,2,3)", expected: 6, source: "pinned", note: "pinned: 1+2+3 = 6." },
  {
    formula: "SUM(A1:A2)",
    cells: { A1: "hello", A2: 5 },
    expected: 5,
    source: "Excel",
    note: 'SUM function, Remarks: "SUM will ignore text values and give you the sum of just the numeric values" — A1 is non-numeric text, ignored.',
  },
  { formula: "SUM(1/0,2)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- SUMIF ---------------------------------------------------------------
  {
    formula: 'SUMIF(A2:A5,">160000",B2:B5)',
    cells: { A2: 100000, A3: 200000, A4: 300000, A5: 400000, B2: 7000, B3: 14000, B4: 21000, B5: 28000 },
    expected: 63000,
    source: "Excel",
    note: "SUMIF function, example 1 (property values/commissions), row 1.",
  },
  {
    formula: "SUMIF(A2:A5,300000,B2:B5)",
    cells: { A2: 100000, A3: 200000, A4: 300000, A5: 400000, B2: 7000, B3: 14000, B4: 21000, B5: 28000 },
    expected: 21000,
    source: "Excel",
    note: "SUMIF function, example 1, row 3.",
  },
  {
    formula: 'SUMIF(A2:A7,"Fruits",C2:C7)',
    cells: { A2: "Vegetables", A3: "Vegetables", A4: "Fruits", A6: "Vegetables", A7: "Fruits", C2: 2300, C3: 5500, C4: 800, C6: 4200, C7: 1200 },
    expected: 2000,
    source: "Excel",
    note: "SUMIF function, example 2 (vegetable/fruit sales), row 1.",
  },
  {
    formula: 'SUMIF(A1,"x",B1,C1)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SUMIF takes at most 3 arguments; a wrong argument count is #VALUE! (SPEC-120 \"Errors\").",
  },

  // --- SUMIFS ---------------------------------------------------------------
  {
    formula: 'SUMIFS(A2:A9,B2:B9,"=A*",C2:C9,"Tom")',
    cells: {
      A2: 5, A3: 4, A4: 15, A5: 3, A6: 22, A7: 12, A8: 10, A9: 33,
      B2: "Apples", B3: "Apples", B4: "Artichokes", B5: "Artichokes", B6: "Bananas", B7: "Bananas", B8: "Carrots", B9: "Carrots",
      C2: "Tom", C3: "Sarah", C4: "Tom", C5: "Sarah", C6: "Tom", C7: "Sarah", C8: "Tom", C9: "Sarah",
    },
    expected: 20,
    source: "Excel",
    note: "SUMIFS function, example 1, row 1.",
  },
  {
    formula: 'SUMIFS(A2:A9,B2:B9,"<>Bananas",C2:C9,"Tom")',
    cells: {
      A2: 5, A3: 4, A4: 15, A5: 3, A6: 22, A7: 12, A8: 10, A9: 33,
      B2: "Apples", B3: "Apples", B4: "Artichokes", B5: "Artichokes", B6: "Bananas", B7: "Bananas", B8: "Carrots", B9: "Carrots",
      C2: "Tom", C3: "Sarah", C4: "Tom", C5: "Sarah", C6: "Tom", C7: "Sarah", C8: "Tom", C9: "Sarah",
    },
    expected: 30,
    source: "Excel",
    note: "SUMIFS function, example 1, row 2.",
  },
  {
    formula: "SUMIFS(A2:A9,B2:B9)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SUMIFS needs at least 3 arguments; a wrong argument count is #VALUE! (SPEC-120 \"Errors\").",
  },

  // --- SUMPRODUCT ---------------------------------------------------------------
  { formula: "SUMPRODUCT({1,2,3},{4,5,6})", expected: 32, source: "pinned", note: "pinned: 1*4+2*5+3*6 = 4+10+18 = 32." },
  {
    formula: "SUMPRODUCT(A1:A2,B1:B2)",
    cells: { A1: 2, A2: "x", B1: 3, B2: 4 },
    expected: 6,
    source: "Excel",
    note: 'SUMPRODUCT function, Remarks: "SUMPRODUCT treats non-numeric array entries as if they were zeros" — 2*3 + 0*4 = 6.',
  },
  {
    formula: "SUMPRODUCT({1,2},{1,2,3})",
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'SUMPRODUCT function, Remarks: "The array arguments must have the same dimensions. If they do not, SUMPRODUCT returns the #VALUE! error value."',
  },

  // --- PRODUCT ---------------------------------------------------------------
  { formula: "PRODUCT(A2:A4)", cells: { A2: 5, A3: 15, A4: 30 }, expected: 2250, source: "Excel", note: "PRODUCT function, example, row 1." },
  { formula: "PRODUCT(A2:A4,2)", cells: { A2: 5, A3: 15, A4: 30 }, expected: 4500, source: "Excel", note: "PRODUCT function, example, row 2." },
  { formula: "PRODUCT(1/0,2)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ROUND ---------------------------------------------------------------
  { formula: "ROUND(2.15,1)", expected: 2.2, digits: 2, source: "Excel", note: "ROUND function, example, row 1." },
  { formula: "ROUND(-1.475,2)", expected: -1.48, digits: 3, source: "Excel", note: "ROUND function, example, row 3." },
  { formula: "ROUND(21.5,-1)", expected: 20, digits: 2, source: "Excel", note: "ROUND function, example, row 4." },
  { formula: "ROUND(1,2,3)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: ROUND takes at most 2 arguments; a wrong argument count is #VALUE!." },

  // --- ROUNDUP ---------------------------------------------------------------
  { formula: "ROUNDUP(3.2,0)", expected: 4, digits: 1, source: "Excel", note: "ROUNDUP function, example, row 1." },
  // eslint-disable-next-line approx-constant -- the result as Microsoft's page prints it, compared at its digits
  { formula: "ROUNDUP(3.14159,3)", expected: 3.142, digits: 4, source: "Excel", note: "ROUNDUP function, example, row 3." },
  { formula: "ROUNDUP(-3.14159,1)", expected: -3.2, digits: 2, source: "Excel", note: "ROUNDUP function, example, row 4." },
  { formula: "ROUNDUP(1,2,3)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: ROUNDUP takes at most 2 arguments; a wrong argument count is #VALUE!." },

  // --- ROUNDDOWN ---------------------------------------------------------------
  { formula: "ROUNDDOWN(3.2,0)", expected: 3, digits: 1, source: "Excel", note: "ROUNDDOWN function, example, row 1." },
  // eslint-disable-next-line approx-constant -- the result as Microsoft's page prints it, compared at its digits
  { formula: "ROUNDDOWN(3.14159,3)", expected: 3.141, digits: 4, source: "Excel", note: "ROUNDDOWN function, example, row 3." },
  { formula: "ROUNDDOWN(-3.14159,1)", expected: -3.1, digits: 2, source: "Excel", note: "ROUNDDOWN function, example, row 4." },
  { formula: "ROUNDDOWN(1,2,3)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: ROUNDDOWN takes at most 2 arguments; a wrong argument count is #VALUE!." },

  // --- INT ---------------------------------------------------------------
  { formula: "INT(8.9)", expected: 8, digits: 1, source: "Excel", note: "INT function, example, row 1." },
  { formula: "INT(-8.9)", expected: -9, digits: 1, source: "Excel", note: "INT function, example, row 2." },
  { formula: "INT(19.5)", expected: 19, source: "pinned", note: "pinned: INT's example (A2 = 19.5) floors to 19." },
  { formula: "INT(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- TRUNC ---------------------------------------------------------------
  { formula: "TRUNC(8.9)", expected: 8, digits: 1, source: "Excel", note: "TRUNC function, example, row 1." },
  { formula: "TRUNC(-8.9)", expected: -8, digits: 1, source: "Excel", note: "TRUNC function, example, row 2." },
  {
    formula: "TRUNC(-4.3)",
    expected: -4,
    digits: 1,
    source: "Excel",
    note: 'TRUNC function, Remarks: "TRUNC(-4.3) returns -4, but INT(-4.3) returns -5."',
  },
  { formula: "TRUNC(1,2,3)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: TRUNC takes at most 2 arguments; a wrong argument count is #VALUE!." },

  // --- MOD ---------------------------------------------------------------
  { formula: "MOD(3,2)", expected: 1, source: "Excel", note: "MOD function, example, row 1." },
  { formula: "MOD(-3,2)", expected: 1, source: "Excel", note: "MOD function, example, row 2." },
  { formula: "MOD(3,-2)", expected: -1, source: "Excel", note: "MOD function, example, row 3." },
  { formula: "MOD(3,0)", expected: { error: "#DIV/0!" }, source: "Excel", note: 'MOD function, Remarks: "If divisor is 0, MOD returns the #DIV/0! error value."' },

  // --- ABS ---------------------------------------------------------------
  { formula: "ABS(2)", expected: 2, source: "Excel", note: "ABS function, example, row 1." },
  { formula: "ABS(-2)", expected: 2, source: "Excel", note: "ABS function, example, row 2." },
  { formula: "ABS(A2)", cells: { A2: -4 }, expected: 4, source: "Excel", note: "ABS function, example, row 3." },
  { formula: "ABS(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: ABS takes exactly 1 argument; a wrong argument count is #VALUE!." },

  // --- SIGN ---------------------------------------------------------------
  { formula: "SIGN(10)", expected: 1, source: "Excel", note: "SIGN function, example, row 1." },
  { formula: "SIGN(4-4)", expected: 0, source: "Excel", note: "SIGN function, example, row 2." },
  { formula: "SIGN(-0.00001)", expected: -1, source: "Excel", note: "SIGN function, example, row 3." },
  { formula: "SIGN(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- SQRT ---------------------------------------------------------------
  { formula: "SQRT(16)", expected: 4, source: "Excel", note: "SQRT function, example, row 1." },
  { formula: "SQRT(-16)", expected: { error: "#NUM!" }, source: "Excel", note: "SQRT function, example, row 2 (A2 = -16)." },
  { formula: "SQRT(ABS(-16))", expected: 4, source: "Excel", note: "SQRT function, example, row 3." },

  // --- POWER ---------------------------------------------------------------
  { formula: "POWER(5,2)", expected: 25, source: "Excel", note: "POWER function, example, row 1." },
  { formula: "POWER(98.6,3.2)", expected: 2401077.222, digits: 10, source: "Excel", note: "POWER function, example, row 2." },
  { formula: "POWER(4,5/4)", expected: 5.656854249, digits: 10, source: "Excel", note: "POWER function, example, row 3." },
  { formula: "POWER(2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: POWER needs 2 arguments; a wrong argument count is #VALUE!." },

  // --- EXP ---------------------------------------------------------------
  // eslint-disable-next-line approx-constant -- the result as Microsoft's page prints it, compared at its digits
  { formula: "EXP(1)", expected: 2.71828183, digits: 9, source: "Excel", note: "EXP function, example, row 1." },
  { formula: "EXP(2)", expected: 7.3890561, digits: 8, source: "Excel", note: "EXP function, example, row 2." },
  { formula: "EXP(0)", expected: 1, source: "pinned", note: "pinned: e^0 = 1." },
  { formula: "EXP(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: EXP takes exactly 1 argument; a wrong argument count is #VALUE!." },

  // --- LN ---------------------------------------------------------------
  { formula: "LN(86)", expected: 4.4543473, digits: 8, source: "Excel", note: "LN function, example, row 1." },
  { formula: "LN(2.7182818)", expected: 1, digits: 1, source: "Excel", note: "LN function, example, row 2." },
  { formula: "LN(EXP(3))", expected: 3, digits: 1, source: "Excel", note: "LN function, example, row 3." },
  { formula: "LN(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- LOG ---------------------------------------------------------------
  { formula: "LOG(10)", expected: 1, digits: 1, source: "Excel", note: "LOG function, example, row 1." },
  { formula: "LOG(8,2)", expected: 3, digits: 1, source: "Excel", note: "LOG function, example, row 2." },
  { formula: "LOG(86,2.7182818)", expected: 4.4543473, digits: 8, source: "Excel", note: "LOG function, example, row 3." },
  { formula: "LOG(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- LOG10 ---------------------------------------------------------------
  { formula: "LOG10(86)", expected: 1.9345, digits: 5, source: "Excel", note: "LOG10 function, example, row 1." },
  { formula: "LOG10(10)", expected: 1, digits: 1, source: "Excel", note: "LOG10 function, example, row 2." },
  { formula: "LOG10(100000)", expected: 5, digits: 1, source: "Excel", note: "LOG10 function, example, row 3." },
  { formula: "LOG10(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- PI ---------------------------------------------------------------
  // eslint-disable-next-line approx-constant -- the result as Microsoft's page prints it, compared at its digits
  { formula: "PI()", expected: 3.141592654, digits: 10, source: "Excel", note: "PI function, example, row 1." },
  { formula: "PI(1)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: PI takes no arguments; a wrong argument count is #VALUE! (e.g. PI(1))." },
  { formula: "PI(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: PI takes no arguments; a wrong argument count is #VALUE!." },

  // --- CEILING ---------------------------------------------------------------
  { formula: "CEILING(2.5,1)", expected: 3, source: "Excel", note: "CEILING function, example, row 1." },
  { formula: "CEILING(-2.5,-2)", expected: -4, source: "Excel", note: "CEILING function, example, row 2." },
  { formula: "CEILING(-2.5,2)", expected: -2, source: "Excel", note: "CEILING function, example, row 3." },
  {
    formula: "CEILING(2.5,-1)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: SPEC-120's CEILING/FLOOR sign rule — a positive number with negative significance is #NUM! (the mirror of FLOOR's documented different-signs error).",
  },

  // --- FLOOR ---------------------------------------------------------------
  { formula: "FLOOR(3.7,2)", expected: 2, source: "Excel", note: "FLOOR function, example, row 1." },
  { formula: "FLOOR(-2.5,-2)", expected: -2, source: "Excel", note: "FLOOR function, example, row 2." },
  { formula: "FLOOR(1.58,0.1)", expected: 1.5, digits: 2, source: "Excel", note: "FLOOR function, example, row 4." },
  {
    formula: "FLOOR(2.5,-2)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "FLOOR function, example, row 3: \"because 2.5 and -2 have different signs.\"",
  },

  // --- MROUND ---------------------------------------------------------------
  { formula: "MROUND(10,3)", expected: 9, source: "Excel", note: "MROUND function, example, row 1." },
  { formula: "MROUND(-10,-3)", expected: -9, source: "Excel", note: "MROUND function, example, row 2." },
  { formula: "MROUND(1.3,0.2)", expected: 1.4, digits: 2, source: "Excel", note: "MROUND function, example, row 3." },
  {
    formula: "MROUND(5,-2)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'MROUND function, example, row 4: "because -2 and 5 have different signs."',
  },

  // --- RAND ---------------------------------------------------------------
  { formula: "RAND()", expected: 0.5, source: "pinned", note: "pinned: SPEC-127's conformance clock draws every RAND() as exactly 0.5." },
  { formula: "RAND(1)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: RAND takes no arguments; a wrong argument count is #VALUE!." },
  { formula: "RAND(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: RAND takes no arguments; a wrong argument count is #VALUE!." },

  // --- RANDBETWEEN ---------------------------------------------------------------
  {
    formula: "RANDBETWEEN(1,10)",
    expected: 6,
    source: "pinned",
    note: "pinned: SPEC-120's formula with random()=0.5 — ceil(1)+floor(0.5*(floor(10)-ceil(1)+1)) = 1+floor(0.5*10) = 1+5 = 6.",
  },
  {
    formula: "RANDBETWEEN(1,1)",
    expected: 1,
    source: "pinned",
    note: "pinned: ceil(1)+floor(0.5*(floor(1)-ceil(1)+1)) = 1+floor(0.5*1) = 1+0 = 1.",
  },
  {
    formula: "RANDBETWEEN(-5,5)",
    expected: 0,
    source: "pinned",
    note: "pinned: ceil(-5)+floor(0.5*(floor(5)-ceil(-5)+1)) = -5+floor(0.5*11) = -5+5 = 0.",
  },
  { formula: "RANDBETWEEN(5,1)", expected: { error: "#NUM!" }, source: "pinned", note: "pinned: bottom > top is #NUM! (SPEC-120)." },

  // --- SUMSQ ---------------------------------------------------------------
  { formula: "SUMSQ(3,4)", expected: 25, source: "Excel", note: "SUMSQ function, example: 3^2+4^2 = 9+16 = 25." },
  { formula: "SUMSQ(2)", expected: 4, source: "pinned", note: "pinned: 2^2 = 4." },
  { formula: "SUMSQ(1,2,3)", expected: 14, source: "pinned", note: "pinned: 1^2+2^2+3^2 = 1+4+9 = 14." },
  { formula: "SUMSQ(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- GCD ---------------------------------------------------------------
  { formula: "GCD(5,2)", expected: 1, source: "Excel", note: "GCD function, example, row 1." },
  { formula: "GCD(24,36)", expected: 12, source: "Excel", note: "GCD function, example, row 2." },
  { formula: "GCD(5,0)", expected: 5, source: "Excel", note: "GCD function, example, row 4." },
  {
    formula: "GCD(9007199254740992,2)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: SPEC-127's pin — a GCD operand ≥ 2^53 (9007199254740992) is #NUM!.",
  },

  // --- LCM ---------------------------------------------------------------
  { formula: "LCM(5,2)", expected: 10, source: "Excel", note: "LCM function, example, row 1." },
  { formula: "LCM(24,36)", expected: 72, source: "Excel", note: "LCM function, example, row 2." },
  { formula: "LCM(4,6)", expected: 12, source: "pinned", note: "pinned: gcd(4,6)=2, lcm = 4*6/2 = 12." },
  {
    formula: "LCM(9007199254740992,2)",
    expected: { error: "#NUM!" },
    source: "pinned",
    note: "pinned: SPEC-127's pin — an LCM operand ≥ 2^53 (9007199254740992) is #NUM!.",
  },

  // --- FACT ---------------------------------------------------------------
  { formula: "FACT(5)", expected: 120, source: "Excel", note: "FACT function, example, row 1: 1*2*3*4*5." },
  { formula: "FACT(0)", expected: 1, source: "Excel", note: "FACT function, example, row 3." },
  { formula: "FACT(-1)", expected: { error: "#NUM!" }, source: "Excel", note: "FACT function, example, row 4." },
  { formula: "FACT(171)", expected: { error: "#NUM!" }, source: "pinned", note: "pinned: SPEC-127's pin — FACT above 170 is #NUM!." },

  // --- COMBIN ---------------------------------------------------------------
  { formula: "COMBIN(8,2)", expected: 28, source: "Excel", note: "COMBIN function, example: possible two-person teams from 8 candidates." },
  { formula: "COMBIN(4,0)", expected: 1, source: "pinned", note: "pinned: choosing 0 of 4 has exactly 1 way." },
  { formula: "COMBIN(10,10)", expected: 1, source: "pinned", note: "pinned: choosing all 10 of 10 has exactly 1 way." },
  {
    formula: "COMBIN(2,5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'COMBIN function, Remarks: "If number < 0, number_chosen < 0, or number < number_chosen, COMBIN returns the #NUM! error value."',
  },

  // --- PERMUT ---------------------------------------------------------------
  { formula: "PERMUT(A2,A3)", cells: { A2: 100, A3: 3 }, expected: 970200, source: "Excel", note: "PERMUT function, example, row 1." },
  { formula: "PERMUT(3,2)", expected: 6, source: "Excel", note: "PERMUT function, example, row 2." },
  { formula: "PERMUT(5,0)", expected: 1, source: "pinned", note: "pinned: permutations of 0 chosen from 5 is 1." },
  {
    formula: "PERMUT(2,5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERMUT function, Remarks: "If number < number_chosen, PERMUT returns the #NUM! error value."',
  },

  // --- QUOTIENT ---------------------------------------------------------------
  { formula: "QUOTIENT(5,2)", expected: 2, source: "Excel", note: "QUOTIENT function, example, row 1." },
  { formula: "QUOTIENT(4.5,3.1)", expected: 1, source: "Excel", note: "QUOTIENT function, example, row 2." },
  { formula: "QUOTIENT(-10,3)", expected: -3, source: "Excel", note: "QUOTIENT function, example, row 3." },
  {
    formula: 'QUOTIENT("a",2)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'QUOTIENT function, Remarks: "If either argument is nonnumeric, QUOTIENT returns the #VALUE! error value."',
  },

  // --- EVEN ---------------------------------------------------------------
  { formula: "EVEN(1.5)", expected: 2, source: "Excel", note: "EVEN function, example, row 1." },
  { formula: "EVEN(3)", expected: 4, source: "Excel", note: "EVEN function, example, row 2." },
  { formula: "EVEN(-1)", expected: -2, source: "Excel", note: "EVEN function, example, row 4 — demonstrates rounding away from zero." },
  {
    formula: 'EVEN("x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'EVEN function, Remarks: "If number is nonnumeric, EVEN returns the #VALUE! error value."',
  },

  // --- ODD ---------------------------------------------------------------
  { formula: "ODD(1.5)", expected: 3, source: "Excel", note: "ODD function, example, row 1." },
  { formula: "ODD(3)", expected: 3, source: "Excel", note: "ODD function, example, row 2." },
  { formula: "ODD(-2)", expected: -3, source: "Excel", note: "ODD function, example, row 5 — demonstrates rounding away from zero." },
  {
    formula: 'ODD("x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'ODD function, Remarks: "If number is nonnumeric, ODD returns the #VALUE! error value."',
  },

  // --- SIN ---------------------------------------------------------------
  // (Microsoft's own "=SIN(PI())" example is skipped: PI() is not exactly
  // the mathematical pi in floating point, so the printed near-zero result
  // is not reliably re-transcribable digit-for-digit — see the report.)
  { formula: "SIN(PI()/2)", expected: 1, source: "pinned", note: "pinned: sin(π/2) = 1 exactly." },
  { formula: "SIN(RADIANS(30))", expected: 0.5, digits: 1, source: "Excel", note: "SIN function, example, row 4." },
  { formula: "SIN(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- COS ---------------------------------------------------------------
  { formula: "COS(1.047)", expected: 0.5001711, digits: 7, source: "Excel", note: "COS function, example, row 1." },
  { formula: "COS(RADIANS(60))", expected: 0.5, digits: 1, source: "Excel", note: "COS function, example, row 3." },
  { formula: "COS(0)", expected: 1, source: "pinned", note: "pinned: cos(0) = 1 exactly." },
  { formula: "COS(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- TAN ---------------------------------------------------------------
  { formula: "TAN(0.785)", expected: 0.9992, digits: 5, source: "Excel", note: "TAN function, example, row 1." },
  { formula: "TAN(RADIANS(45))", expected: 1, digits: 1, source: "Excel", note: "TAN function, example, row 3." },
  { formula: "TAN(0)", expected: 0, source: "pinned", note: "pinned: tan(0) = 0 exactly." },
  { formula: "TAN(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ASIN ---------------------------------------------------------------
  { formula: "ASIN(-0.5)", expected: -0.523598776, digits: 9, source: "Excel", note: "ASIN function, example, row 1 (-pi/6)." },
  { formula: "ASIN(1)", expected: 1.570796327, digits: 10, source: "pinned", note: "pinned: asin(1) = π/2 = 1.570796327…" },
  {
    formula: "ASIN(1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "ASIN function: \"Number must be between -1 and 1\" — 1.5 is outside the domain.",
  },

  // --- ACOS ---------------------------------------------------------------
  { formula: "ACOS(-0.5)", expected: 2.094395102, digits: 10, source: "Excel", note: "ACOS function, example, row 1 (2*pi/3)." },
  { formula: "ACOS(1)", expected: 0, source: "pinned", note: "pinned: acos(1) = 0 exactly." },
  {
    formula: "ACOS(1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "ACOS function: \"Number must be between -1 and 1\" — 1.5 is outside the domain.",
  },

  // --- ATAN ---------------------------------------------------------------
  { formula: "ATAN(1)", expected: 0.785398163, digits: 9, source: "Excel", note: "ATAN function, example, row 1 (pi/4)." },
  { formula: "ATAN(0)", expected: 0, source: "pinned", note: "pinned: atan(0) = 0 exactly." },
  { formula: "ATAN(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ATAN2 ---------------------------------------------------------------
  { formula: "ATAN2(1,1)", expected: 0.785398163, digits: 9, source: "Excel", note: "ATAN2 function, example, row 1 (pi/4)." },
  { formula: "ATAN2(-1,-1)", expected: -2.35619449, digits: 9, source: "Excel", note: "ATAN2 function, example, row 2 (-3*pi/4)." },
  {
    formula: "ATAN2(0,0)",
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'ATAN2 function, Remarks: "If both x_num and y_num are 0, ATAN2 returns the #DIV/0! error value."',
  },

  // --- SINH ---------------------------------------------------------------
  // Microsoft's SINH example formula multiplies SINH(...) by a constant (its
  // root is "*", not SINH), so it cannot root a row here; these are hand
  // derived instead (its page carries an example table, but not one usable
  // at SINH's own root — see the report).
  { formula: "SINH(0)", expected: 0, source: "pinned", note: "pinned: sinh(0) = 0 exactly." },
  { formula: "SINH(1)", expected: 1.1752011936438014, source: "pinned", note: "pinned: (e^1 - e^-1)/2 = (2.718281828459045 - 0.36787944117144233)/2 = 1.1752011936438014." },
  { formula: "SINH(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- COSH ---------------------------------------------------------------
  { formula: "COSH(4)", expected: 27.308233, digits: 8, source: "Excel", note: "COSH function, example, row 1." },
  { formula: "COSH(EXP(1))", expected: 7.6101251, digits: 8, source: "Excel", note: "COSH function, example, row 2." },
  { formula: "COSH(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- TANH ---------------------------------------------------------------
  { formula: "TANH(-2)", expected: -0.964028, digits: 6, source: "Excel", note: "TANH function, example, row 1." },
  { formula: "TANH(0)", expected: 0, source: "Excel", note: "TANH function, example, row 2." },
  { formula: "TANH(0.5)", expected: 0.462117, digits: 6, source: "Excel", note: "TANH function, example, row 3." },
  { formula: "TANH(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ASINH ---------------------------------------------------------------
  { formula: "ASINH(-2.5)", expected: -1.647231146, digits: 10, source: "Excel", note: "ASINH function, example, row 1." },
  { formula: "ASINH(10)", expected: 2.99822295, digits: 9, source: "Excel", note: "ASINH function, example, row 2." },
  { formula: "ASINH(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ACOSH ---------------------------------------------------------------
  { formula: "ACOSH(1)", expected: 0, source: "Excel", note: "ACOSH function, example, row 1." },
  { formula: "ACOSH(10)", expected: 2.9932228, digits: 8, source: "Excel", note: "ACOSH function, example, row 2." },
  {
    formula: "ACOSH(0.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "ACOSH function: \"Number must be a real value ≥ 1\" — 0.5 is outside the domain.",
  },

  // --- ATANH ---------------------------------------------------------------
  { formula: "ATANH(0.76159416)", expected: 1.00000001, digits: 9, source: "Excel", note: "ATANH function, example, row 1." },
  { formula: "ATANH(-0.1)", expected: -0.100335348, digits: 9, source: "Excel", note: "ATANH function, example, row 2." },
  {
    formula: "ATANH(1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "ATANH function: \"Number between -1 and 1, excluding the endpoints\" — 1 is an endpoint.",
  },

  // --- COT ---------------------------------------------------------------
  { formula: "COT(30)", expected: -0.156, digits: 3, source: "Excel", note: "COT function, example, row 1." },
  { formula: "COT(A5)", cells: { A5: 45 }, expected: 0.617, digits: 3, source: "Excel", note: "COT function, example, row 2." },
  { formula: "COT(0)", expected: { error: "#DIV/0!" }, source: "Excel", note: 'COT function, Remarks: "COT(0) returns the #DIV/0! error value."' },
  {
    formula: "COT(134217728)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "COT function, Remarks: \"The absolute value of Number must be less than 2^27\" — 134217728 = 2^27.",
  },

  // --- CSC ---------------------------------------------------------------
  { formula: "CSC(15)", expected: 1.538, digits: 4, source: "Excel", note: "CSC function, example." },
  {
    formula: "CSC(134217728)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "CSC function, Remarks: \"The absolute value of Number must be less than 2^27\" — 134217728 = 2^27.",
  },
  { formula: "CSC(0)", expected: { error: "#DIV/0!" }, source: "pinned", note: 'pinned: CSC function, Remarks: "CSC(n) equals 1/SIN(n)" — 1/sin(0) = 1/0.' },

  // --- SEC ---------------------------------------------------------------
  { formula: "SEC(45)", expected: 1.90359, digits: 6, source: "Excel", note: "SEC function, example, row 1." },
  { formula: "SEC(30)", expected: 6.48292, digits: 6, source: "Excel", note: "SEC function, example, row 2." },
  {
    formula: "SEC(134217728)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "SEC function, Remarks: \"The absolute value of number must be less than 2^27\" — 134217728 = 2^27.",
  },

  // --- ACOT ---------------------------------------------------------------
  { formula: "ACOT(2)", expected: 0.4636, digits: 4, source: "Excel", note: "ACOT function, example." },
  { formula: "ACOT(1)", expected: 0.7853981633974483, source: "pinned", note: "pinned: acot(1) = atan(1/1) = π/4 = 0.7853981633974483." },
  {
    formula: 'ACOT("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'ACOT function, Remarks: "If Number is a non-numeric value, ACOT returns the #VALUE! error value."',
  },

  // --- DEGREES ---------------------------------------------------------------
  { formula: "DEGREES(PI())", expected: 180, source: "pinned", note: "pinned: 180° = π radians, exactly 180." },
  { formula: "DEGREES(ATAN(1))", expected: 45, source: "Excel", note: "ATAN function, example, row 3 (DEGREES of pi/4)." },
  { formula: "DEGREES(ATAN2(-1,-1))", expected: -135, source: "Excel", note: "ATAN2 function, example, row 4 (DEGREES of -3*pi/4)." },
  { formula: "DEGREES(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- RADIANS ---------------------------------------------------------------
  { formula: "RADIANS(270)", expected: 4.712389, digits: 7, source: "Excel", note: "RADIANS function, example." },
  { formula: "RADIANS(90)", expected: 1.570796327, digits: 10, source: "pinned", note: "pinned: 90° = π/2 = 1.570796327…" },
  { formula: "RADIANS(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- CEILING.MATH ---------------------------------------------------------------
  { formula: "CEILING.MATH(24.3,5)", expected: 25, source: "Excel", note: "CEILING.MATH function, example, row 1." },
  { formula: "CEILING.MATH(-8.1,2)", expected: -8, source: "Excel", note: "CEILING.MATH function, example, row 3 (default mode: toward 0)." },
  {
    formula: "CEILING.MATH(-5.5,2,-1)",
    expected: -6,
    source: "Excel",
    note: "CEILING.MATH function, example, row 4 (nonzero mode reverses direction: away from 0).",
  },
  { formula: "CEILING.MATH(1/0,2)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- FLOOR.MATH ---------------------------------------------------------------
  { formula: "FLOOR.MATH(24.3,5)", expected: 20, source: "Excel", note: "FLOOR.MATH function, example, row 1." },
  { formula: "FLOOR.MATH(-8.1,2)", expected: -10, source: "Excel", note: "FLOOR.MATH function, example, row 3 (default mode: away from 0)." },
  {
    formula: "FLOOR.MATH(-5.5,2,-1)",
    expected: -4,
    source: "Excel",
    note: "FLOOR.MATH function, example, row 4 (nonzero mode reverses direction: toward 0).",
  },
  { formula: "FLOOR.MATH(1/0,2)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- CEILING.PRECISE ---------------------------------------------------------------
  { formula: "CEILING.PRECISE(4.3)", expected: 5, source: "Excel", note: "CEILING.PRECISE function, example, row 1." },
  { formula: "CEILING.PRECISE(-4.3,2)", expected: -4, source: "Excel", note: "CEILING.PRECISE function, example, row 5." },
  { formula: "CEILING.PRECISE(4.3,-2)", expected: 6, source: "Excel", note: "CEILING.PRECISE function, example, row 4 (sign of significance ignored)." },
  { formula: "CEILING.PRECISE(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- FLOOR.PRECISE ---------------------------------------------------------------
  { formula: "FLOOR.PRECISE(-3.2,-1)", expected: -4, source: "Excel", note: "FLOOR.PRECISE function, example, row 1." },
  { formula: "FLOOR.PRECISE(3.2,1)", expected: 3, source: "Excel", note: "FLOOR.PRECISE function, example, row 2." },
  { formula: "FLOOR.PRECISE(3.2)", expected: 3, source: "Excel", note: "FLOOR.PRECISE function, example, row 5 (significance defaults to 1)." },
  { formula: "FLOOR.PRECISE(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ISO.CEILING ---------------------------------------------------------------
  { formula: "ISO.CEILING(4.3)", expected: 5, source: "Excel", note: "ISO.CEILING function, example, row 2." },
  { formula: "ISO.CEILING(-4.3,2)", expected: -4, source: "Excel", note: "ISO.CEILING function, example, row 6." },
  { formula: "ISO.CEILING(4.3,-2)", expected: 6, source: "Excel", note: "ISO.CEILING function, example, row 5 (sign of significance ignored)." },
  { formula: "ISO.CEILING(1/0)", expected: { error: "#DIV/0!" }, source: "pinned", note: "pinned: an error argument propagates (SPEC-120 \"Errors\")." },

  // --- ROMAN ---------------------------------------------------------------
  { formula: "ROMAN(499,0)", expected: "CDXCIX", source: "Excel", note: "ROMAN function, example, row 1 (classic style)." },
  {
    formula: "ROMAN(3999,0)",
    expected: "MMMCMXCIX",
    source: "pinned",
    note: "pinned: 3999 = 3000(MMM)+900(CM)+90(XC)+9(IX), the classic form's largest value.",
  },
  {
    formula: "ROMAN(4000,0)",
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'ROMAN function, Remarks: "If number is greater than 3999, the #VALUE! error value is returned."',
  },
  {
    formula: "ROMAN(499,1)",
    expected: { error: "#NAME?" },
    source: "pinned",
    note: "pinned: this library computes only form 0/TRUE (SPEC-127 \"Function pins\"); form 1 throws UnsupportedCaseError, and an uncomputable cell with no cached value is #NAME? (SPEC-127).",
  },

  // --- ARABIC ---------------------------------------------------------------
  { formula: 'ARABIC("LVII")', expected: 57, source: "Excel", note: "ARABIC function, example, row 1." },
  { formula: "ARABIC(B1)", cells: { B1: "mcmxii" }, expected: 1912, source: "Excel", note: "ARABIC function, example, row 2 (case-insensitive)." },
  { formula: 'ARABIC("")', expected: 0, source: "Excel", note: "ARABIC function, Remarks: an empty string returns 0." },
  {
    formula: 'ARABIC("ABC")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'ARABIC function, Remarks: "Values that return a #VALUE! error value include... text that is not a valid Roman numeral."',
  },

  // --- BASE ---------------------------------------------------------------
  { formula: "BASE(7,2)", expected: "111", source: "Excel", note: "BASE function, example, row 1." },
  { formula: "BASE(100,16)", expected: "64", source: "Excel", note: "BASE function, example, row 2." },
  { formula: "BASE(15,2,10)", expected: "0000001111", source: "Excel", note: "BASE function, example, row 3 (padded to min_length 10)." },
  {
    formula: "BASE(7,37)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'BASE function, Remarks: "If Number, Radix, or Min_length are outside the minimum or maximum constraints, BASE returns the #NUM! error value" — radix must be ≤ 36.',
  },

  // --- DECIMAL ---------------------------------------------------------------
  { formula: 'DECIMAL("FF",16)', expected: 255, source: "Excel", note: "DECIMAL function, example 1." },
  { formula: "DECIMAL(111,2)", expected: 7, source: "Excel", note: "DECIMAL function, example 2." },
  { formula: 'DECIMAL("zap",36)', expected: 45745, source: "Excel", note: "DECIMAL function, example 3." },
  {
    formula: 'DECIMAL("10",37)',
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'DECIMAL function: "Radix must be greater than or equal to 2... and less than or equal to 36" — 37 is out of range.',
  },

  // --- MMULT ---------------------------------------------------------------
  // Microsoft's MMULT example data is embedded in images, not machine-
  // transcribable text (see the report); these are hand-built and verified.
  {
    formula: "MMULT(A1:B2,D1:E2)",
    cells: { A1: 1, B1: 2, A2: 3, B2: 4, D1: 5, E1: 6, D2: 7, E2: 8 },
    array: "G1:H2",
    expected: 19,
    expectedArray: [
      [19, 22],
      [43, 50],
    ],
    source: "pinned",
    note: "pinned: [[1,2],[3,4]] * [[5,6],[7,8]] = [[1*5+2*7,1*6+2*8],[3*5+4*7,3*6+4*8]] = [[19,22],[43,50]].",
  },
  { formula: "MMULT({1,2},{3;4})", expected: 11, source: "pinned", note: "pinned: a 1x2 times a 2x1 gives a 1x1 result: 1*3+2*4 = 11." },
  {
    formula: "MMULT({1,2},{3,4})",
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'MMULT function, Remarks: "The number of columns in array1 must be the same as the number of rows in array2" — 2 columns vs. 1 row.',
  },

  // --- SUBTOTAL ---------------------------------------------------------------
  {
    formula: "SUBTOTAL(9,A2:A5)",
    cells: { A2: 120, A3: 10, A4: 150, A5: 23 },
    expected: 303,
    source: "Excel",
    note: "SUBTOTAL function, example, row 1 (sum, function_num 9).",
  },
  {
    formula: "SUBTOTAL(1,A2:A5)",
    cells: { A2: 120, A3: 10, A4: 150, A5: 23 },
    expected: 75.75,
    source: "Excel",
    note: "SUBTOTAL function, example, row 2 (average, function_num 1).",
  },
  {
    formula: "SUBTOTAL(4,A2:A5)",
    cells: { A2: 120, A3: 10, A4: 150, A5: 23 },
    expected: 150,
    source: "pinned",
    note: "pinned: MAX(120,10,150,23) = 150 (function_num 4, SPEC-127's SUBTOTAL table).",
  },
  {
    formula: "SUBTOTAL(109,A2:A5)",
    cells: { A2: 120, A3: 10, A4: 150, A5: 23 },
    expected: 303,
    source: "pinned",
    note: "pinned: 109 (SUM excluding hidden rows) behaves as 9 when no row is hidden.",
  },
  {
    formula: "SUBTOTAL(12,A2:A5)",
    cells: { A2: 120, A3: 10, A4: 150, A5: 23 },
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-127 \"SUBTOTAL and AGGREGATE\": any function_num outside 1-11/101-111 is #VALUE! for SUBTOTAL.",
  },

  // --- AGGREGATE ---------------------------------------------------------------
  // Hand-built, visibility-free dataset (SPEC-127's brief sanctions pinned
  // rows on visible data here — Microsoft's own worked example renders
  // inconsistently across transcriptions, see the report).
  {
    formula: "AGGREGATE(9,6,A1:A4)",
    cells: { A1: 10, A2: { error: "#DIV/0!" }, A3: 20, A4: 30 },
    expected: 60,
    source: "pinned",
    note: "pinned: SUM (function_num 9) ignoring errors (option 6): 10+20+30 = 60.",
  },
  {
    formula: "AGGREGATE(4,6,A1:A4)",
    cells: { A1: 10, A2: { error: "#DIV/0!" }, A3: 20, A4: 30 },
    expected: 30,
    source: "pinned",
    note: "pinned: MAX (function_num 4) ignoring errors (option 6): max(10,20,30) = 30.",
  },
  {
    formula: "AGGREGATE(15,6,A1:A4)",
    cells: { A1: 10, A2: { error: "#DIV/0!" }, A3: 20, A4: 30 },
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "AGGREGATE function, example: SMALL (function_num 15, array form) needs a k argument; omitting it is #VALUE!.",
  },
  {
    formula: "AGGREGATE(20,6,A1:A4)",
    cells: { A1: 10, A2: { error: "#DIV/0!" }, A3: 20, A4: 30 },
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-127 \"SUBTOTAL and AGGREGATE\": AGGREGATE supports function_num 1-19 only; 20 is #VALUE!.",
  },
];
