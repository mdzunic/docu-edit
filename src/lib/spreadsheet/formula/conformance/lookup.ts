// SPEC-127 "Conformance tables": the Lookup & Reference conformance rows.
// Sources: Microsoft's function reference pages
// (support.microsoft.com/en-us/excel/functions/<name>-function) and their
// Remarks, and SPEC-120/127's stated rules (spec.md "Function pins",
// "Errors") for pinned rows. VLOOKUP, XLOOKUP, TRANSPOSE and HYPERLINK's own
// example sections are image-only or depend on an external workbook/URL in
// every row (skipped per the brief), so those are pinned from the functions'
// documented argument rules; SINGLE has no function page at all (it is
// Excel's implicit-intersection operator "@" in file form) and is pinned
// entirely from SPEC-120's stated rule. Never this engine's own output.
import type { ConformanceRow } from "./row";

export const LOOKUP_ROWS: ConformanceRow[] = [
  // --- VLOOKUP ---------------------------------------------------------------
  // VLOOKUP's own example tables are image-only (no recoverable cell data);
  // rows pinned from the page's stated Remarks (exact/approximate match,
  // smaller-than-smallest, wildcard support).
  {
    formula: "VLOOKUP(2,A1:B3,2,FALSE)",
    cells: { A1: 1, B1: "a", A2: 2, B2: "b", A3: 3, B3: "c" },
    expected: "b",
    source: "pinned",
    note: 'pinned: VLOOKUP function Remarks — exact match (FALSE) "searches for the exact value in the first column."',
  },
  {
    formula: "VLOOKUP(15,A1:B3,2,TRUE)",
    cells: { A1: 1, B1: "low", A2: 10, B2: "mid", A3: 20, B3: "high" },
    expected: "mid",
    source: "pinned",
    note: "pinned: VLOOKUP function Remarks — approximate match (TRUE) finds the largest value <= 15, which is 10.",
  },
  {
    formula: "VLOOKUP(5,A1:B2,2,TRUE)",
    cells: { A1: 10, B1: "x", A2: 20, B2: "y" },
    expected: { error: "#N/A" },
    source: "Excel",
    note: "VLOOKUP function Remarks: lookup_value smaller than the smallest value in the first column returns #N/A.",
  },
  {
    formula: 'VLOOKUP("ban*",A1:B2,2,FALSE)',
    cells: { A1: "apple", B1: 1, A2: "banana", B2: 2 },
    expected: 2,
    source: "Excel",
    note: "VLOOKUP function Remarks: with range_lookup FALSE and a text lookup_value, wildcards ? and * are supported.",
  },

  // --- HLOOKUP -----------------------------------------------------------------
  {
    formula: "HLOOKUP(\"Axles\", A1:C4, 2, TRUE)",
    cells: { A1: "Axles", B1: "Bearings", C1: "Bolts", A2: 4, B2: 4, C2: 9, A3: 5, B3: 7, C3: 10, A4: 6, B4: 8, C4: 11 },
    expected: 4,
    source: "Excel",
    note: "HLOOKUP function example 1: Axles in row 1, value from row 2 same column.",
  },
  {
    formula: "HLOOKUP(\"Bearings\", A1:C4, 3, FALSE)",
    cells: { A1: "Axles", B1: "Bearings", C1: "Bolts", A2: 4, B2: 4, C2: 9, A3: 5, B3: 7, C3: 10, A4: 6, B4: 8, C4: 11 },
    expected: 7,
    source: "Excel",
    note: "HLOOKUP function example 2: exact match on Bearings, value from row 3 same column.",
  },
  {
    formula: "HLOOKUP(\"B\", A1:C4, 3, TRUE)",
    cells: { A1: "Axles", B1: "Bearings", C1: "Bolts", A2: 4, B2: 4, C2: 9, A3: 5, B3: 7, C3: 10, A4: 6, B4: 8, C4: 11 },
    expected: 5,
    source: "Excel",
    note: 'HLOOKUP function example 3: no exact match for "B", largest value less than "B" is "Axles".',
  },
  {
    formula: 'HLOOKUP("Bolts", A1:C4, 4)',
    cells: { A1: "Axles", B1: "Bearings", C1: "Bolts", A2: 4, B2: 4, C2: 9, A3: 5, B3: 7, C3: 10, A4: 6, B4: 8, C4: 11 },
    expected: 11,
    source: "Excel",
    note: "HLOOKUP function example 4: range_lookup omitted (defaults TRUE), Bolts found, value from row 4.",
  },
  {
    formula: 'HLOOKUP(3, {1,2,3;"a","b","c";"d","e","f"}, 2, TRUE)',
    expected: "c",
    source: "Excel",
    note: "HLOOKUP function example 5: looks up 3 in an array constant's first row, returns row 2 of the same column.",
  },
  {
    formula: 'HLOOKUP("Aaa", A1:C4, 2, TRUE)',
    cells: { A1: "Axles", B1: "Bearings", C1: "Bolts", A2: 4, B2: 4, C2: 9, A3: 5, B3: 7, C3: 10, A4: 6, B4: 8, C4: 11 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'HLOOKUP function Remarks: "If lookup_value is smaller than the smallest value in the first row ... HLOOKUP returns the #N/A error value."',
  },

  // --- INDEX -------------------------------------------------------------------
  {
    formula: "INDEX(A2:B3,2,2)",
    cells: { A2: "Apples", B2: "Lemons", A3: "Bananas", B3: "Pears" },
    expected: "Pears",
    source: "Excel",
    note: "INDEX function, array form example: row 2, column 2 of A2:B3.",
  },
  {
    formula: "INDEX(A2:B3,2,1)",
    cells: { A2: "Apples", B2: "Lemons", A3: "Bananas", B3: "Pears" },
    expected: "Bananas",
    source: "Excel",
    note: "INDEX function, array form example: row 2, column 1 of A2:B3.",
  },
  {
    formula: "INDEX(A2:C6, 2, 3)",
    cells: {
      A1: "Fruit", B1: "Price", C1: "Count",
      A2: "Apples", B2: 0.69, C2: 40,
      A3: "Bananas", B3: 0.34, C3: 38,
      A4: "Lemons", B4: 0.55, C4: 15,
      A5: "Oranges", B5: 0.25, C5: 25,
      A6: "Pears", B6: 0.59, C6: 40,
    },
    expected: 38,
    source: "Excel",
    note: "INDEX function, reference form example: intersection of row 2 and column 3 of A2:C6 (Bananas' count).",
  },
  {
    formula: "INDEX({1,2;3,4},0,2)",
    expected: 2,
    source: "pinned",
    note: "pinned: row_num 0 selects the whole of column 2 of an array constant (no reference), the array {2;4}; a plain cell's array result is its top-left element (SPEC-120 \"Implicit intersection\"), so 2.",
  },

  // --- MATCH -------------------------------------------------------------------
  {
    formula: "MATCH(39,B2:B5,1)",
    cells: { A1: "Product", B1: "Count", A2: "Bananas", B2: 25, A3: "Oranges", B3: 38, A4: "Apples", B4: 40, A5: "Pears", B5: 41 },
    expected: 2,
    source: "Excel",
    note: "MATCH function example 1: no exact match, position of the next lowest value (38).",
  },
  {
    formula: "MATCH(41,B2:B5,0)",
    cells: { A1: "Product", B1: "Count", A2: "Bananas", B2: 25, A3: "Oranges", B3: 38, A4: "Apples", B4: 40, A5: "Pears", B5: 41 },
    expected: 4,
    source: "Excel",
    note: "MATCH function example 2: exact match, position of 41.",
  },
  {
    formula: "MATCH(40,B2:B5,-1)",
    cells: { A1: "Product", B1: "Count", A2: "Bananas", B2: 25, A3: "Oranges", B3: 38, A4: "Apples", B4: 40, A5: "Pears", B5: 41 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: "MATCH function example 3: match_type -1 expects descending order; B2:B5 is ascending, so #N/A.",
  },

  // --- ROW ---------------------------------------------------------------------
  { formula: "ROW(C10)", expected: 10, source: "Excel", note: "ROW function example: the row of an explicit reference, independent of the host cell." },
  { formula: "ROW()", expected: 100, source: "pinned", note: "pinned: SPEC-127 — every conformance formula is entered at Z100; ROW() with no argument is the host row, 100." },
  { formula: "ROW(Z1)", expected: 1, source: "pinned", note: "pinned: ROW function — the row of an explicit reference, Z1 is row 1 (generalising the page's ROW(C10) example)." },
  { formula: "ROW(A1,A2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- COLUMN ------------------------------------------------------------------
  { formula: "COLUMN(D10)", expected: 4, source: "Excel", note: "COLUMN function: \"the formula =COLUMN(D10) returns 4, because column D is the fourth column.\"" },
  { formula: "COLUMN()", expected: 26, source: "pinned", note: "pinned: every conformance formula is entered at Z100; COLUMN() with no argument is the host column, Z = 26." },
  { formula: "COLUMN(A1:C1)", expected: 1, source: "pinned", note: "pinned: COLUMN function — not entered as an array formula, a multi-column reference gives its first column." },
  { formula: "COLUMN(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- ROWS --------------------------------------------------------------------
  { formula: "ROWS(C1:E4)", expected: 4, source: "Excel", note: "ROWS function example: the number of rows in the reference C1:E4." },
  { formula: "ROWS({1,2,3;4,5,6})", expected: 2, source: "Excel", note: "ROWS function example: the number of rows in a 2-row array constant." },
  { formula: "ROWS(A1)", expected: 1, source: "pinned", note: "pinned: ROWS function — a single cell is a one-row reference." },
  { formula: "ROWS()", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (min 1) is #VALUE!." },

  // --- COLUMNS -----------------------------------------------------------------
  { formula: "COLUMNS(C1:E4)", expected: 3, source: "Excel", note: "COLUMNS function example: the number of columns in the reference C1:E4." },
  {
    formula: "COLUMNS({1,2,3;4,5,6})",
    expected: 3,
    source: "Excel",
    note: "COLUMNS function example: two 3-column rows, {1,2,3} and {4,5,6}.",
  },
  { formula: "COLUMNS(A1)", expected: 1, source: "pinned", note: "pinned: COLUMNS function — a single cell is a one-column reference." },
  { formula: "COLUMNS()", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (min 1) is #VALUE!." },

  // --- OFFSET ------------------------------------------------------------------
  {
    formula: "OFFSET(D3,3,-2,1,1)",
    cells: { B6: 4 },
    expected: 4,
    source: "Excel",
    note: 'OFFSET function example: "Displays the value in cell B6 (4)" — D3 down 3 rows, left 2 columns is B6; OFFSET\'s root shows the referenced cell\'s value.',
  },
  {
    formula: "OFFSET(D3, -3, -3)",
    expected: { error: "#REF!" },
    source: "Excel",
    note: "OFFSET function example: \"Returns an error, because the reference is to a non-existent range on the worksheet\" (row 3-3=0 is before row 1).",
  },
  { formula: "OFFSET(A1,0,0)", cells: { A1: 9 }, expected: 9, source: "pinned", note: "pinned: OFFSET function — a zero row/column offset with no height/width is the starting reference itself." },

  // --- INDIRECT ----------------------------------------------------------------
  {
    formula: "INDIRECT(A2)",
    cells: { A2: "B2", B2: 1.333 },
    expected: 1.333,
    source: "Excel",
    note: "INDIRECT function example: A2 holds the text B2, referring to B2 (1.333).",
  },
  {
    formula: "INDIRECT(A3)",
    cells: { A3: "B3", B3: 45 },
    expected: 45,
    source: "Excel",
    note: "INDIRECT function example: A3 holds the text B3, referring to B3 (45).",
  },
  {
    formula: 'INDIRECT("B"&A5)',
    cells: { A5: 5, B5: 62 },
    expected: 62,
    source: "Excel",
    note: 'INDIRECT function example: "B"&A5 combines to "B5", referring to B5 (62).',
  },
  { formula: "INDIRECT(A1,A2,A3)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 2) is #VALUE!." },

  // --- SINGLE ------------------------------------------------------------------
  // SINGLE has no Microsoft function page — it is Excel's implicit
  // intersection operator "@" in file form. Rows pinned entirely from
  // SPEC-120's stated rule (spec.md "Function pins" / SPEC-127 brief).
  { formula: "SINGLE(A1)", cells: { A1: 5 }, expected: 5, source: "pinned", note: "pinned: SPEC-120 — SINGLE of a single cell is its value." },
  {
    formula: "SINGLE(A1:A200)",
    cells: { A100: 42 },
    expected: 42,
    source: "pinned",
    note: "pinned: SPEC-120 — SINGLE of a one-column range takes the cell in the formula's row (100), here A100.",
  },
  {
    formula: "SINGLE(A1:A50)",
    cells: { A1: 1 },
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — a one-column range that does not include the formula's row (100) is #VALUE!.",
  },
  { formula: "SINGLE({1,2;3,4})", expected: 1, source: "pinned", note: "pinned: SPEC-120 — SINGLE of an array is its top-left element." },

  // --- CHOOSE ------------------------------------------------------------------
  {
    formula: "CHOOSE(2,A2,A3,A4,A5)",
    cells: { A2: "1st", A3: "2nd", A4: "3rd", A5: "Finished" },
    expected: "2nd",
    source: "Excel",
    note: "CHOOSE function example: index 2 selects the second value argument (A3).",
  },
  {
    formula: "CHOOSE(4,B2,B3,B4,B5)",
    cells: { B2: "Nails", B3: "Screws", B4: "Nuts", B5: "Bolts" },
    expected: "Bolts",
    source: "Excel",
    note: "CHOOSE function example: index 4 selects the fourth value argument (B5).",
  },
  {
    formula: 'CHOOSE(3,"Wide",115,"world",8)',
    expected: "world",
    source: "Excel",
    note: "CHOOSE function example: index 3 selects the third value argument.",
  },
  {
    formula: "CHOOSE(0,1,2)",
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'CHOOSE function Remarks: "If index_num is less than 1 ... CHOOSE returns the #VALUE! error value."',
  },

  // --- XLOOKUP -----------------------------------------------------------------
  // XLOOKUP's Microsoft examples that spill (return arrays into several
  // cells) cannot be expressed; the single-cell examples on the page are
  // image-only with no recoverable literal values, so rows are pinned from
  // SPEC-127's explicit pins and XLOOKUP's basic documented behaviour.
  {
    formula: "XLOOKUP(\"b\",A1:A3,B1:B3)",
    cells: { A1: "a", A2: "b", A3: "c", B1: 10, B2: 20, B3: 30 },
    expected: 20,
    source: "pinned",
    note: "pinned: XLOOKUP function — default exact match of the lookup value against lookup_array, returning the same position of return_array.",
  },
  {
    formula: "XLOOKUP(\"b\",A1:A3,B1:B2)",
    cells: { A1: "a", A2: "b", A3: "c", B1: 10, B2: 20 },
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — XLOOKUP with mismatched lookup_array/return_array sizes is #VALUE!.",
  },
  {
    formula: "XLOOKUP(\"z\",A1:A3,B1:B3)",
    cells: { A1: "a", A2: "b", A3: "c", B1: 10, B2: 20, B3: 30 },
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — not found without if_not_found is #N/A.",
  },

  // --- XMATCH ------------------------------------------------------------------
  {
    formula: "XMATCH(4,{5,4,3,2,1})",
    expected: 2,
    source: "Excel",
    note: "XMATCH function example: exact match (default), position of 4 in the array.",
  },
  {
    formula: "XMATCH(4.5,{5,4,3,2,1},1)",
    expected: 1,
    source: "Excel",
    note: "XMATCH function example: match_mode 1, exact match or next largest item (5, at position 1).",
  },
  {
    formula: "XMATCH(99,{5,4,3,2,1})",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: XMATCH has no if_not_found for exact match mode 0; not found is #N/A, as MATCH documents for the analogous case.",
  },

  // --- LOOKUP ------------------------------------------------------------------
  {
    formula: "LOOKUP(4.19, A2:A6, B2:B6)",
    cells: { A2: 4.14, A3: 4.19, A4: 5.17, A5: 5.77, A6: 6.39, B2: "red", B3: "orange", B4: "yellow", B5: "green", B6: "blue" },
    expected: "orange",
    source: "Excel",
    note: "LOOKUP function, vector form example 1: exact match on 4.19.",
  },
  {
    formula: "LOOKUP(5.75, A2:A6, B2:B6)",
    cells: { A2: 4.14, A3: 4.19, A4: 5.17, A5: 5.77, A6: 6.39, B2: "red", B3: "orange", B4: "yellow", B5: "green", B6: "blue" },
    expected: "yellow",
    source: "Excel",
    note: "LOOKUP function, vector form example 2: no exact match, nearest smaller value 5.17.",
  },
  {
    formula: "LOOKUP(7.66, A2:A6, B2:B6)",
    cells: { A2: 4.14, A3: 4.19, A4: 5.17, A5: 5.77, A6: 6.39, B2: "red", B3: "orange", B4: "yellow", B5: "green", B6: "blue" },
    expected: "blue",
    source: "Excel",
    note: "LOOKUP function, vector form example 3: no exact match, nearest smaller value 6.39 (the largest).",
  },
  {
    formula: "LOOKUP(0, A2:A6, B2:B6)",
    cells: { A2: 4.14, A3: 4.19, A4: 5.17, A5: 5.77, A6: 6.39, B2: "red", B3: "orange", B4: "yellow", B5: "green", B6: "blue" },
    expected: { error: "#N/A" },
    source: "Excel",
    note: "LOOKUP function, vector form example 4: 0 is smaller than the smallest lookup value (4.14), #N/A.",
  },

  // --- ADDRESS -----------------------------------------------------------------
  { formula: "ADDRESS(2,3)", expected: "$C$2", source: "Excel", note: "ADDRESS function example: abs_num omitted (1), absolute reference." },
  { formula: "ADDRESS(2,3,2)", expected: "C$2", source: "Excel", note: "ADDRESS function example: abs_num 2, absolute row, relative column." },
  {
    formula: "ADDRESS(2,3,2,FALSE)",
    expected: "R2C[3]",
    source: "Excel",
    note: "ADDRESS function example: abs_num 2 in R1C1 style, absolute row (R2), relative column (C[3]).",
  },
  {
    formula: 'ADDRESS(2,3,1,FALSE,"EXCEL SHEET")',
    expected: "'EXCEL SHEET'!R2C3",
    source: "Excel",
    note: "ADDRESS function example: a sheet_text containing a space is quoted.",
  },
  { formula: "ADDRESS(2,3,3)", expected: "$C2", source: "pinned", note: "pinned: ADDRESS function's abs_num table — 3 is relative row, absolute column." },
  { formula: "ADDRESS(2,3,4)", expected: "C2", source: "pinned", note: "pinned: ADDRESS function's abs_num table — 4 is fully relative." },
  {
    formula: "ADDRESS(2,3,3,FALSE)",
    expected: "R[2]C3",
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — abs_num 3 in R1C1 style is R[2]C3.",
  },
  {
    formula: "ADDRESS(2,3,4,FALSE)",
    expected: "R[2]C[3]",
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — abs_num 4 (relative) in R1C1 style is R[2]C[3].",
  },
  {
    formula: "ADDRESS(2,3,1,FALSE)",
    expected: "R2C3",
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — abs_num 1 (absolute) in R1C1 style is R2C3.",
  },
  { formula: "ADDRESS(2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (min 2) is #VALUE!." },

  // --- AREAS -------------------------------------------------------------------
  { formula: "AREAS(B2:D4)", expected: 1, source: "Excel", note: "AREAS function example: a single range is one area." },
  {
    formula: "AREAS((B2:D4,E5,F6:I9))",
    expected: 3,
    source: "Excel",
    note: 'AREAS function example: a parenthesised union reference of three ranges, per SPEC-127 Function pins ("AREAS counts the areas of a union reference").',
  },
  { formula: "AREAS(A1)", expected: 1, source: "pinned", note: "pinned: AREAS function — a single cell reference is one area." },
  { formula: "AREAS(B2:D4,B5)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- TRANSPOSE ---------------------------------------------------------------
  // TRANSPOSE's own page states only the concept ("=TRANSPOSE(A1:B4) ...
  // arranges horizontally") with no recoverable numeric table; rows pinned
  // from the matrix-transpose definition, entered as legacy array formulas.
  {
    formula: "TRANSPOSE(A1:B2)",
    cells: { A1: 1, B1: 3, A2: 2, B2: 4 },
    array: "D1:E2",
    expected: 1,
    expectedArray: [
      [1, 2],
      [3, 4],
    ],
    source: "pinned",
    note: "pinned: TRANSPOSE function — transposing the 2x2 matrix [[1,3],[2,4]] gives [[1,2],[3,4]]; the anchor D1 is 1.",
  },
  {
    formula: "TRANSPOSE(A1:C1)",
    cells: { A1: 5, B1: 6, C1: 7 },
    array: "A3:A5",
    expected: 5,
    expectedArray: [[5], [6], [7]],
    source: "pinned",
    note: "pinned: TRANSPOSE function — a 1x3 row transposes to a 3x1 column; the anchor A3 is 5.",
  },
  { formula: "TRANSPOSE(A1:B2,A1)", cells: { A1: 1 }, expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- HYPERLINK ---------------------------------------------------------------
  // Every row of HYPERLINK's own example table links to an external
  // workbook, URL or UNC path (skipped per the brief); rows pinned from
  // SPEC-127's explicit pin (friendly_name if given, else link).
  {
    formula: 'HYPERLINK("http://example.com")',
    expected: "http://example.com",
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — HYPERLINK(link,[friendly_name]) shows link when friendly_name is omitted.",
  },
  {
    formula: 'HYPERLINK("http://example.com","Click here")',
    expected: "Click here",
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — HYPERLINK shows friendly_name when given.",
  },
  {
    formula: 'HYPERLINK("http://example.com",A1)',
    cells: { A1: "Report" },
    expected: "Report",
    source: "pinned",
    note: "pinned: SPEC-127 Function pins — friendly_name may come from a cell reference.",
  },
  { formula: "HYPERLINK()", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (min 1) is #VALUE!." },

  // --- FORMULATEXT -------------------------------------------------------------
  {
    formula: "FORMULATEXT(A2)",
    cells: { A2: "=TODAY()" },
    expected: "=TODAY()",
    source: "Excel",
    note: "FORMULATEXT function example: returns the formula in A2, =TODAY(), as text.",
  },
  {
    formula: "FORMULATEXT(A1)",
    cells: { A1: 5 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: "FORMULATEXT function Remarks: the referenced cell does not contain a formula, so #N/A.",
  },
  {
    formula: "FORMULATEXT(A1:B1)",
    cells: { A1: "=1+1", B1: "=2+2" },
    expected: "=1+1",
    source: "Excel",
    note: "FORMULATEXT function Remarks: a multi-cell reference returns the formula of its upper-leftmost cell.",
  },
];
