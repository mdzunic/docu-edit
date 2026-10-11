// SPEC-127 "Conformance tables": the Information conformance rows.
// Sources: Microsoft's function reference pages
// (support.microsoft.com/en-us/excel/functions/<name>-function); ISBLANK,
// ISNUMBER, ISTEXT, ISERROR, ISERR, ISNA, ISLOGICAL and ISREF are documented
// together on the combined "is-functions" page (individual per-function
// URLs 404). SHEET, SHEETS, CELL and INFO's own examples rely on a
// multi-sheet workbook, named 3D references or info types this workbook
// cannot reproduce (this conformance workbook always has exactly one sheet,
// "Sheet1"), so those four, plus every row for TYPE's 1/4 codes and
// ERROR.TYPE's 9-14 codes beyond Microsoft's documented 1-8, are pinned from
// SPEC-127's "Function pins" and "Function metadata". Never this engine's
// own output.
import type { ConformanceRow } from "./row";

export const INFORMATION_ROWS: ConformanceRow[] = [
  // --- ISBLANK -------------------------------------------------------------
  { formula: "ISBLANK(A2)", cells: { A2: "Gold" }, expected: false, source: "Excel", note: "IS functions page, Example 2: A2 holds text, not blank." },
  { formula: "ISBLANK(A1)", expected: true, source: "pinned", note: "pinned: A1 is a genuinely empty cell." },
  {
    formula: "ISBLANK(A1)",
    cells: { A1: '=""' },
    expected: false,
    source: "pinned",
    note: 'pinned: SPEC-120 — ISBLANK is TRUE only for an empty cell; a formula result of "" is not blank.',
  },
  { formula: "ISBLANK(A1,A2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISNUMBER --------------------------------------------------------------
  { formula: "ISNUMBER(4)", expected: true, source: "Excel", note: "IS functions page, Example 1: 4 is a number." },
  { formula: "ISNUMBER(A5)", cells: { A5: 330.92 }, expected: true, source: "Excel", note: "IS functions page, Example 2: A5 = 330.92, a number." },
  { formula: 'ISNUMBER("4")', expected: false, source: "pinned", note: 'pinned: IS functions page Remarks — values in quotation marks are text, not converted.' },
  { formula: "ISNUMBER(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISTEXT ----------------------------------------------------------------
  { formula: "ISTEXT(A3)", cells: { A3: "Region1" }, expected: true, source: "Excel", note: "IS functions page, Example 2: A3 holds the text Region1." },
  { formula: "ISTEXT(4)", expected: false, source: "pinned", note: "pinned: 4 is a number, not text." },
  { formula: "ISTEXT(TRUE)", expected: false, source: "pinned", note: "pinned: TRUE is a logical value, not text." },
  { formula: "ISTEXT(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISERROR -----------------------------------------------------------------
  {
    formula: "ISERROR(A4)",
    cells: { A4: { error: "#REF!" } },
    expected: true,
    source: "Excel",
    note: "IS functions page, Example 2: A4 is #REF!, any error.",
  },
  { formula: "ISERROR(5)", expected: false, source: "pinned", note: "pinned: 5 is not an error." },
  { formula: "ISERROR(A1)", cells: { A1: { error: "#N/A" } }, expected: true, source: "pinned", note: "pinned: ISERROR catches every error code, including #N/A." },
  { formula: "ISERROR(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISERR ---------------------------------------------------------------
  {
    formula: "ISERR(A6)",
    cells: { A6: { error: "#N/A" } },
    expected: false,
    source: "Excel",
    note: "IS functions page, Example 2: ISERR excludes #N/A specifically.",
  },
  { formula: "ISERR(A1)", cells: { A1: { error: "#REF!" } }, expected: true, source: "pinned", note: "pinned: ISERR catches errors other than #N/A." },
  { formula: "ISERR(5)", expected: false, source: "pinned", note: "pinned: 5 is not an error." },
  { formula: "ISERR(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISNA ----------------------------------------------------------------
  { formula: "ISNA(A4)", cells: { A4: { error: "#REF!" } }, expected: false, source: "Excel", note: "IS functions page, Example 2: A4 is #REF!, not #N/A." },
  { formula: "ISNA(A6)", cells: { A6: { error: "#N/A" } }, expected: true, source: "Excel", note: "IS functions page, Example 2: A6 is #N/A." },
  { formula: "ISNA(5)", expected: false, source: "pinned", note: "pinned: 5 is not an error." },
  { formula: "ISNA(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISLOGICAL -------------------------------------------------------------
  { formula: "ISLOGICAL(TRUE)", expected: true, source: "Excel", note: "IS functions page, Example 1: TRUE is a logical value." },
  { formula: 'ISLOGICAL("TRUE")', expected: false, source: "Excel", note: 'IS functions page, Example 1: the text "TRUE" is not a logical value.' },
  { formula: "ISLOGICAL(1)", expected: false, source: "pinned", note: "pinned: 1 is a number, not a logical value." },
  { formula: "ISLOGICAL(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- NA ------------------------------------------------------------------
  { formula: "NA()", expected: { error: "#N/A" }, source: "Excel", note: 'NA function — "Returns the error value #N/A."' },
  { formula: "NA(1)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-127 — NA() takes no arguments; NA(1) is #VALUE!." },
  { formula: "NA(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-127 — NA() takes no arguments; NA(1,2) is #VALUE!." },

  // --- ISEVEN ----------------------------------------------------------------
  { formula: "ISEVEN(-1)", expected: false, source: "Excel", note: "ISEVEN function example: -1 is odd." },
  { formula: "ISEVEN(2.5)", expected: true, source: "Excel", note: "ISEVEN function example: 2.5 truncates to 2, even." },
  { formula: "ISEVEN(0)", expected: true, source: "Excel", note: "ISEVEN function example: zero is considered even." },
  { formula: 'ISEVEN("x")', expected: { error: "#VALUE!" }, source: "Excel", note: 'ISEVEN function Remarks: "If number is nonnumeric, ISEVEN returns the #VALUE! error value."' },

  // --- ISODD -----------------------------------------------------------------
  { formula: "ISODD(-1)", expected: true, source: "Excel", note: "ISODD function example: -1 is odd." },
  { formula: "ISODD(2.5)", expected: false, source: "Excel", note: "ISODD function example: 2.5 truncates to 2, even." },
  { formula: "ISODD(5)", expected: true, source: "Excel", note: "ISODD function example: 5 is odd." },
  { formula: 'ISODD("x")', expected: { error: "#VALUE!" }, source: "Excel", note: 'ISODD function Remarks: "If number is nonnumeric, ISODD returns the #VALUE! error value."' },

  // --- ISREF -----------------------------------------------------------------
  { formula: "ISREF(G8)", expected: true, source: "Excel", note: "IS functions page, Example 1: G8 is a valid reference." },
  { formula: "ISREF(XYZ1)", expected: false, source: "Excel", note: "IS functions page, Example 1: XYZ1 names no defined reference." },
  { formula: "ISREF(A1)", cells: { A1: 5 }, expected: true, source: "pinned", note: "pinned: a plain cell reference, whatever its contents, is a reference." },
  { formula: "ISREF(A1,A2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ISFORMULA -------------------------------------------------------------
  { formula: "ISFORMULA(A1)", cells: { A1: "=TODAY()" }, expected: true, source: "Excel", note: "ISFORMULA function example: A1 contains the formula =TODAY()." },
  {
    formula: "ISFORMULA(A1)",
    cells: { A1: "=3/0" },
    expected: true,
    source: "Excel",
    note: "ISFORMULA function example: a formula that errors is still a formula.",
  },
  { formula: "ISFORMULA(A1)", cells: { A1: 7 }, expected: false, source: "Excel", note: "ISFORMULA function example: 7 is a number, not a formula." },
  {
    formula: "ISFORMULA(A1)",
    cells: { A1: "Hello, world!" },
    expected: false,
    source: "Excel",
    note: "ISFORMULA function example: text is not a formula.",
  },
  { formula: "ISFORMULA(A1,A2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- TYPE ------------------------------------------------------------------
  { formula: "TYPE(A2)", cells: { A2: "Smith" }, expected: 2, source: "Excel", note: "TYPE function example: text type is 2." },
  { formula: 'TYPE("Mr. "&A2)', cells: { A2: "Smith" }, expected: 2, source: "Excel", note: "TYPE function example: a concatenated text is still type 2." },
  {
    formula: "TYPE(2+A2)",
    cells: { A2: "Smith" },
    expected: 16,
    source: "Excel",
    note: "TYPE function example: 2+A2 errors #VALUE! (text can't add); TYPE of an error is 16 (TYPE accepts errors).",
  },
  { formula: "TYPE({1,2;3,4})", expected: 64, source: "Excel", note: "TYPE function example: an array constant is type 64." },
  { formula: "TYPE(5)", expected: 1, source: "pinned", note: "pinned: SPEC-127 Function pins — TYPE: 1 number or blank." },
  { formula: "TYPE(TRUE)", expected: 4, source: "pinned", note: "pinned: SPEC-127 Function pins — TYPE: 4 logical." },
  { formula: "TYPE(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ERROR.TYPE --------------------------------------------------------------
  // Microsoft's page documents codes 1-8 explicitly and states "anything
  // else" returns #N/A; it does not enumerate 9-14. Codes 9 and 14 below are
  // pinned from SPEC-127's own extension table (spec.md "Function pins"),
  // since Microsoft's page stops at #GETTING_DATA (8).
  { formula: "ERROR.TYPE(A1)", cells: { A1: { error: "#NULL!" } }, expected: 1, source: "Excel", note: "ERROR.TYPE function: #NULL! is error number 1." },
  { formula: "ERROR.TYPE(A1)", cells: { A1: { error: "#DIV/0!" } }, expected: 2, source: "Excel", note: "ERROR.TYPE function: #DIV/0! is error number 2." },
  { formula: "ERROR.TYPE(A1)", cells: { A1: { error: "#N/A" } }, expected: 7, source: "Excel", note: "ERROR.TYPE function: #N/A is error number 7." },
  {
    formula: "ERROR.TYPE(A1)",
    cells: { A1: { error: "#SPILL!" } },
    expected: 9,
    source: "pinned",
    note: "pinned: SPEC-127 Function pins extends ERROR.TYPE to 9 for #SPILL! — not documented on Microsoft's page, which stops at 8.",
  },
  {
    formula: "ERROR.TYPE(A1)",
    cells: { A1: { error: "#CALC!" } },
    expected: 14,
    source: "pinned",
    note: "pinned: SPEC-127 Function pins extends ERROR.TYPE to 14 for #CALC! — not documented on Microsoft's page, which stops at 8.",
  },
  {
    formula: "ERROR.TYPE(A1)",
    cells: { A1: { error: "#BUSY!" } },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'ERROR.TYPE function: "anything else" (an error code outside the documented list) returns #N/A; #BUSY! is outside both Microsoft\'s and SPEC-127\'s lists.',
  },
  { formula: "ERROR.TYPE(A1,A2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- N ---------------------------------------------------------------------
  { formula: "N(A2)", cells: { A2: 7 }, expected: 7, source: "Excel", note: "N function example: a number passes through." },
  { formula: "N(A3)", cells: { A3: "Even" }, expected: 0, source: "Excel", note: "N function example: text becomes 0." },
  { formula: "N(A4)", cells: { A4: true }, expected: 1, source: "Excel", note: "N function example: TRUE becomes 1." },
  { formula: "N(A5)", cells: { A5: 40650 }, expected: 40650, source: "Excel", note: "N function example: A5 is 4/17/2011 (serial 40650), passes through." },
  { formula: 'N("7")', expected: 0, source: "Excel", note: 'N function example: the text "7" becomes 0.' },
  { formula: "N(A2,A3)", cells: { A2: 7, A3: 2 }, expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- SHEET -------------------------------------------------------------------
  // SHEET's own examples depend on named ranges and tables ("QSalesByRegion",
  // "Table1") or a specific multi-sheet workbook this single-sheet
  // conformance workbook cannot reproduce; pinned from the stated rule for
  // our one sheet, "Sheet1".
  { formula: "SHEET()", expected: 1, source: "pinned", note: "pinned: SHEET function — omitted value returns the sheet containing the formula; this workbook's only sheet is Sheet1 (1)." },
  { formula: 'SHEET("Sheet1")', expected: 1, source: "pinned", note: "pinned: SHEET function — the workbook's one sheet, Sheet1, is sheet number 1." },
  {
    formula: 'SHEET("NotASheet")',
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'SHEET function Remarks: "If the value argument is not a valid sheet name, SHEET returns the #N/A error value."',
  },
  { formula: "SHEET(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- SHEETS ------------------------------------------------------------------
  // SHEETS' own examples use a 3-sheet workbook and a named 3-D reference;
  // pinned from the stated rule for our one-sheet workbook.
  { formula: "SHEETS()", expected: 1, source: "pinned", note: 'pinned: SHEETS function — "with no reference, the total number of sheets in the workbook," here 1.' },
  { formula: "SHEETS(A1)", cells: { A1: 5 }, expected: 1, source: "pinned", note: "pinned: SHEETS function — a reference's sheets, here just Sheet1 (1)." },
  { formula: "SHEETS(A1,A2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- CELL --------------------------------------------------------------------
  // CELL's own Example section is image-only; rows pinned from the
  // info_type meanings the page's Syntax section states, applied to cells
  // this conformance workbook defines. "filename" and the unsupported
  // "width" info type follow SPEC-127's explicit pins.
  { formula: 'CELL("row",A5)', expected: 5, source: "pinned", note: 'pinned: CELL function — "row": row number of the cell, A5 is row 5.' },
  { formula: 'CELL("col",C5)', expected: 3, source: "pinned", note: 'pinned: CELL function — "col": column number of the cell, C is column 3.' },
  { formula: 'CELL("contents",A1)', cells: { A1: 42 }, expected: 42, source: "pinned", note: 'pinned: CELL function — "contents": value of the upper-left cell in reference.' },
  { formula: 'CELL("type",A1)', expected: "b", source: "pinned", note: 'pinned: CELL function — "type": "b" for blank, and A1 is empty.' },
  { formula: 'CELL("type",A1)', cells: { A1: "hi" }, expected: "l", source: "pinned", note: 'pinned: CELL function — "type": "l" for a text label.' },
  { formula: 'CELL("type",A1)', cells: { A1: 5 }, expected: "v", source: "pinned", note: 'pinned: CELL function — "type": "v" for a value (number).' },
  { formula: 'CELL("address",B7)', expected: "$B$7", source: "pinned", note: 'pinned: CELL function — "address": reference of the first cell, as text, absolute.' },
  {
    formula: 'CELL("filename",A1)',
    expected: "[Calc.xlsx]Sheet1",
    source: "pinned",
    note: 'pinned: SPEC-127 Function pins — "filename" is "[<title>.<ext>]<sheet>"; this workbook is Calc.xlsx, sheet Sheet1.',
  },
  { formula: 'CELL("protect",A1)', cells: { A1: 5 }, expected: 0, source: "pinned", note: 'pinned: CELL function — "protect": 0 when the cell is unlocked (the default).' },
  {
    formula: 'CELL("width",A1)',
    expected: { error: "#NAME?" },
    source: "pinned",
    note: 'pinned: SPEC-127 Function pins — "width" is outside CELL\'s supported info types; an uncomputable cell with no cached value is #NAME?.',
  },
  { formula: "CELL()", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (min 1) is #VALUE!." },

  // --- INFO --------------------------------------------------------------------
  {
    formula: 'INFO("numfile")',
    expected: 1,
    source: "pinned",
    note: 'pinned: INFO function — "numfile": number of worksheets in open workbooks; this workbook has one sheet.',
  },
  {
    formula: 'INFO("recalc")',
    expected: "Automatic",
    source: "pinned",
    note: 'pinned: INFO function — "recalc": current recalculation mode; this workbook\'s default is Automatic.',
  },
  {
    formula: 'INFO("system")',
    expected: { error: "#NAME?" },
    source: "pinned",
    note: 'pinned: SPEC-127 Function pins — "system" is outside INFO\'s supported info types (the engine may not read the platform); an uncomputable cell with no cached value is #NAME?.',
  },
  { formula: "INFO()", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (min 1) is #VALUE!." },
];
