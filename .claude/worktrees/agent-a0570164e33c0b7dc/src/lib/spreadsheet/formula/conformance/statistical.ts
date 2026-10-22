// The Statistical and Compatibility conformance rows (SPEC-127 "Conformance
// tables"): SPEC-120's aggregates (AVERAGE … MINIFS), this spec's statistics
// and distributions, and every legacy name Excel files under Compatibility,
// each from its own page.
//
// Sources: each function's Microsoft reference page
// (support.microsoft.com/en-us/excel/functions/<name>-function) — its example
// table for "Excel" rows, and its Remarks for the error rows labelled Excel
// (the smallest input meeting the stated condition) — and, for "pinned" rows,
// hand arithmetic on small data chosen so the result is exact, the derivation
// in each note. Data sits where the page's sheet has it (the example table
// pasted at A1, headers in row 1); dates and times the page shows as such are
// DATE/TIME formula cells. Never this engine's own output, nor any other
// implementation's.
import type { ConformanceRow } from "./row";

export const STATISTICAL_ROWS: ConformanceRow[] = [
  // SPEC-120's aggregates. Pinned error rows use SPEC-120's ArgKind rule (a
  // `range` argument given a non-reference is #VALUE!) or the evaluator's
  // argument-count rule (outside minArgs…maxArgs is #VALUE!; Excel refuses
  // to enter such a call).
  // --- AVERAGE: number1, [number2], … ----------------------------------
  {
    formula: "AVERAGE(A2:A6)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 },
    expected: 11,
    source: "Excel",
    note: "AVERAGE function, example: mean of 10, 7, 9, 27, 2 in A2:A6; Microsoft prints 11.",
  },
  {
    formula: "AVERAGE(A2:A6,5)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 },
    expected: 10,
    source: "Excel",
    note: "AVERAGE function, example: the same five numbers in A2:A6 plus the literal 5; Microsoft prints 10.",
  },
  {
    formula: "AVERAGE(A1)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: A1 is left empty, so AVERAGE has zero numbers to divide its sum by; aggregate.ts's average() returns #DIV/0! for a zero count, the same rule AVERAGEIF and AVERAGEIFS document explicitly.",
  },

  // --- AVERAGEA: value1, [value2], … -----------------------------------
  {
    formula: "AVERAGEA(A2:A6)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 2, A6: "Not available" },
    expected: 5.6,
    source: "Excel",
    note: 'AVERAGEA function, example: mean of 10, 7, 9, 2 and the text "Not available" (counted as 0) in A2:A6; Microsoft prints 5.6.',
  },
  {
    formula: 'AVERAGEA("x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'AVERAGEA function Remarks: "Arguments that are error values or text that cannot be translated into numbers cause errors"; a direct text argument that is not numeric text is #VALUE!.',
  },
  {
    formula: "AVERAGEA(TRUE,FALSE,2)",
    expected: 1,
    source: "pinned",
    note: "pinned: AVERAGEA counts TRUE as 1 and FALSE as 0, so (1+0+2)/3 = 1 exactly.",
  },

  // --- AVERAGEIF: range, criteria, [average_range] ---------------------
  {
    formula: 'AVERAGEIF(B2:B5,"<23000")',
    cells: { A2: 100000, B2: 7000, A3: 200000, B3: 14000, A4: 300000, B4: 21000, A5: 400000, B5: 28000 },
    expected: 14000,
    source: "Excel",
    note: "AVERAGEIF function, example 1: mean of the commissions in B2:B5 under 23000 — 7000, 14000 and 21000, excluding 28000; Microsoft prints 14000.",
  },
  {
    formula: 'AVERAGEIF(A2:A5,"<95000")',
    cells: { A2: 100000, B2: 7000, A3: 200000, B3: 14000, A4: 300000, B4: 21000, A5: 400000, B5: 28000 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: "AVERAGEIF function, example 1: no property value in A2:A5 is under 95000, so there is nothing to average; Microsoft prints #DIV/0! (\"because it tries to divide by 0\").",
  },
  {
    formula: 'AVERAGEIF(A2:A5,">250000",B2:B5)',
    cells: { A2: 100000, B2: 7000, A3: 200000, B3: 14000, A4: 300000, B4: 21000, A5: 400000, B5: 28000 },
    expected: 24500,
    source: "Excel",
    note: "AVERAGEIF function, example 1: mean of the commissions (B2:B5) where the property value (A2:A5) exceeds 250000 — 21000 and 28000; Microsoft prints 24500.",
  },

  // --- AVERAGEIFS: average_range, criteria_range1, criteria1, … --------
  {
    formula: 'AVERAGEIFS(B2:B7,C2:C7,"Bellevue",D2:D7,">2",E2:E7,"Yes")',
    cells: {
      B2: 230000, C2: "Issaquah", D2: 3, E2: "No",
      B3: 197000, C3: "Bellevue", D3: 2, E3: "Yes",
      B4: 345678, C4: "Bellevue", D4: 4, E4: "Yes",
      B5: 321900, C5: "Issaquah", D5: 2, E5: "Yes",
      B6: 450000, C6: "Bellevue", D6: 5, E6: "Yes",
      B7: 395000, C7: "Bellevue", D7: 4, E7: "No",
    },
    expected: 397839,
    source: "Excel",
    note: "AVERAGEIFS function, example 2: mean price (B) of Bellevue listings (C) with more than 2 bedrooms (D) and a garage (E) — 345678 and 450000; Microsoft prints 397839.",
  },
  {
    formula: 'AVERAGEIFS(B2:B7,C2:C7,"Issaquah",D2:D7,"<=3",E2:E7,"No")',
    cells: {
      B2: 230000, C2: "Issaquah", D2: 3, E2: "No",
      B3: 197000, C3: "Bellevue", D3: 2, E3: "Yes",
      B4: 345678, C4: "Bellevue", D4: 4, E4: "Yes",
      B5: 321900, C5: "Issaquah", D5: 2, E5: "Yes",
      B6: 450000, C6: "Bellevue", D6: 5, E6: "Yes",
      B7: 395000, C7: "Bellevue", D7: 4, E7: "No",
    },
    expected: 230000,
    source: "Excel",
    note: "AVERAGEIFS function, example 2: mean price of Issaquah listings with at most 3 bedrooms and no garage — only row 2 (230000) qualifies; Microsoft prints 230000.",
  },
  {
    formula: 'AVERAGEIFS(B2:B7,C2:C7,"Issaquah",D2:D7,">4",E2:E7,"Yes")',
    cells: {
      B2: 230000, C2: "Issaquah", D2: 3, E2: "No",
      B3: 197000, C3: "Bellevue", D3: 2, E3: "Yes",
      B4: 345678, C4: "Bellevue", D4: 4, E4: "Yes",
      B5: 321900, C5: "Issaquah", D5: 2, E5: "Yes",
      B6: 450000, C6: "Bellevue", D6: 5, E6: "Yes",
      B7: 395000, C7: "Bellevue", D7: 4, E7: "No",
    },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'AVERAGEIFS function Remarks: "If there are no cells that meet all the criteria, AVERAGEIFS returns the #DIV/0! error value"; neither Issaquah listing (rows 2, 5) has more than 4 bedrooms.',
  },

  // --- COUNT: value1, [value2], … --------------------------------------
  {
    formula: "COUNT(A2:A6)",
    cells: { A2: "=DATE(2008,12,8)", A3: 19, A4: 22.24, A5: true, A6: { error: "#DIV/0!" } },
    expected: 3,
    source: "Excel",
    note: "COUNT function, example: counts the numeric cells in A2:A6 — a date, 19 and 22.24 — ignoring TRUE and the #DIV/0! error; Microsoft prints 3.",
  },
  {
    formula: "COUNT(A2:A6,2)",
    cells: { A2: "=DATE(2008,12,8)", A3: 19, A4: 22.24, A5: true, A6: { error: "#DIV/0!" } },
    expected: 4,
    source: "Excel",
    note: "COUNT function, example: the same range plus the literal 2, one more than the previous example; Microsoft prints 4.",
  },
  {
    formula: "COUNT()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — a wrong argument count (0, needs at least 1) is #VALUE!.",
  },

  // --- COUNTA: value1, [value2], … -------------------------------------
  {
    formula: "COUNTA(A2:A6)",
    cells: { A2: "=DATE(2008,12,8)", A3: 19, A4: 22.24, A5: true, A6: { error: "#DIV/0!" } },
    expected: 5,
    source: "Excel",
    note: "COUNTA function, example: every one of A2:A6 is non-blank — a date, a number, a decimal, TRUE and an error — so all five count; Microsoft prints 5.",
  },
  {
    formula: "COUNTA()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — a wrong argument count (0, needs at least 1) is #VALUE!.",
  },
  {
    formula: "COUNTA(1,,3)",
    expected: 3,
    source: "pinned",
    note: "pinned: every direct argument counts for COUNTA, even an empty slot, so 1, the empty slot and 3 all count: 3 exactly.",
  },

  // --- COUNTBLANK: range -------------------------------------------------
  {
    formula: "COUNTBLANK(A2:B4)",
    cells: { A2: 6, B2: 27, A3: 4, B3: 34 },
    expected: 2,
    source: "Excel",
    note: "COUNTBLANK function, example: A4 and B4 are the only empty cells in A2:B4; Microsoft prints 2.",
  },
  {
    formula: "COUNTBLANK(A1:A3)",
    cells: { A1: '=""', A2: 0 },
    expected: 2,
    source: "Excel",
    note: "COUNTBLANK function Remarks: cells with formulas that return empty text are also counted, and cells with zero values are not; A1's =\"\" and the never-set A3 count, A2's 0 does not.",
  },
  {
    formula: "COUNTBLANK(5)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: 'pinned: COUNTBLANK\'s range argument given a non-reference is #VALUE! (SPEC-120\'s ArgKind table: "range — anything else → #VALUE!").',
  },

  // --- COUNTIF: range, criteria -------------------------------------------
  {
    formula: 'COUNTIF(A2:A5,"apples")',
    cells: { A2: "apples", B2: 32, A3: "oranges", B3: 54, A4: "peaches", B4: 75, A5: "apples", B5: 86 },
    expected: 2,
    source: "Excel",
    note: 'COUNTIF function, example: counts the two cells in A2:A5 that read "apples"; Microsoft prints 2.',
  },
  {
    formula: 'COUNTIF(B2:B5,">55")',
    cells: { A2: "apples", B2: 32, A3: "oranges", B3: 54, A4: "peaches", B4: 75, A5: "apples", B5: 86 },
    expected: 2,
    source: "Excel",
    note: "COUNTIF function, example: counts the values in B2:B5 greater than 55 — 75 and 86; Microsoft prints 2.",
  },
  {
    formula: 'COUNTIF(A2:A5,"?????es")',
    cells: { A2: "apples", B2: 32, A3: "oranges", B3: 54, A4: "peaches", B4: 75, A5: "apples", B5: 86 },
    expected: 2,
    source: "Excel",
    note: 'COUNTIF function, example: the wildcard pattern matches exactly-seven-character entries ending "es" — "peaches" and "oranges"; Microsoft prints 2.',
  },
  {
    formula: "COUNTIF(5,1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: 'pinned: COUNTIF\'s range argument given a non-reference is #VALUE! (SPEC-120\'s ArgKind table: "range — anything else → #VALUE!").',
  },

  // --- COUNTIFS: criteria_range1, criteria1, … ----------------------------
  {
    formula: 'COUNTIFS(B2:B5,"=Yes",C2:C5,"=Yes")',
    cells: { B2: "Yes", C2: "No", B3: "Yes", C3: "Yes", B4: "Yes", C4: "Yes", B5: "No", C5: "Yes" },
    expected: 2,
    source: "Excel",
    note: "COUNTIFS function, example 1: salespeople who exceeded both their Q1 (B) and Q2 (C) quotas — rows 3 and 4; Microsoft prints 2.",
  },
  {
    formula: 'COUNTIFS(A2:A7,"<6",A2:A7,">1")',
    cells: { A2: 1, A3: 2, A4: 3, A5: 4, A6: 5, A7: 6 },
    expected: 4,
    source: "Excel",
    note: "COUNTIFS function, example 2: counts 1 through 6 in A2:A7 strictly between 1 and 6 — 2, 3, 4 and 5; Microsoft prints 4.",
  },
  {
    formula: 'COUNTIFS(A2:A7,"<5",B2:B7,"<5/3/2011")',
    cells: {
      A2: 1, B2: "=DATE(2011,5,1)",
      A3: 2, B3: "=DATE(2011,5,2)",
      A4: 3, B4: "=DATE(2011,5,3)",
      A5: 4, B5: "=DATE(2011,5,4)",
      A6: 5, B6: "=DATE(2011,5,5)",
      A7: 6, B7: "=DATE(2011,5,6)",
    },
    expected: 2,
    source: "Excel",
    note: "COUNTIFS function, example 2: rows where A is under 5 and B's date is before 5/3/2011 — rows 2 and 3 (5/1 and 5/2); Microsoft prints 2.",
  },
  {
    formula: "COUNTIFS(5,1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: 'pinned: COUNTIFS\'s criteria_range argument given a non-reference is #VALUE! (SPEC-120\'s ArgKind table: "range — anything else → #VALUE!").',
  },

  // --- MAX: number1, [number2], … --------------------------------------
  {
    formula: "MAX(A2:A6)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 },
    expected: 27,
    source: "Excel",
    note: "MAX function, example: the largest of 10, 7, 9, 27, 2 in A2:A6; Microsoft prints 27.",
  },
  {
    formula: "MAX(A2:A6,30)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 },
    expected: 30,
    source: "Excel",
    note: "MAX function, example: the same five numbers plus the literal 30, which becomes the largest; Microsoft prints 30.",
  },
  {
    formula: 'MAX("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'MAX function Remarks: "Arguments that are error values or text that cannot be translated into numbers cause errors"; a direct, non-numeric text argument is #VALUE!.',
  },

  // --- MAXA: value1, [value2], … ----------------------------------------
  {
    formula: "MAXA(A2:A6)",
    cells: { A2: 0, A3: 0.2, A4: 0.5, A5: 0.4, A6: true },
    expected: 1,
    source: "Excel",
    note: "MAXA function, example: the largest of 0, 0.2, 0.5, 0.4 and TRUE (counted as 1) in A2:A6; Microsoft prints 1.",
  },
  {
    formula: 'MAXA("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'MAXA function Remarks: "Arguments that are error values or text that cannot be translated into numbers cause errors"; a direct, non-numeric text argument is #VALUE!.',
  },
  {
    formula: "MAXA(FALSE,0.3,TRUE)",
    expected: 1,
    source: "pinned",
    note: "pinned: MAXA counts FALSE as 0 and TRUE as 1, so the largest of 0, 0.3 and 1 is 1 exactly.",
  },

  // --- MAXIFS: max_range, criteria_range1, criteria1, … -------------------
  {
    formula: "MAXIFS(A2:A7,B2:B7,1)",
    cells: { A2: 89, B2: 1, A3: 93, B3: 2, A4: 96, B4: 2, A5: 85, B5: 3, A6: 91, B6: 1, A7: 88, B7: 1 },
    expected: 91,
    source: "Excel",
    note: "MAXIFS function, example 1: the largest grade (A) among the rows weighted 1 (B) — 89, 91 and 88; Microsoft prints 91.",
  },
  {
    formula: 'MAXIFS(A2:A7,B2:B7,"b",D2:D7,">100")',
    cells: { A2: 10, B2: "b", D2: 100, A3: 1, B3: "a", D3: 100, A4: 100, B4: "a", D4: 200, A5: 1, B5: "b", D5: 300, A6: 1, B6: "a", D6: 100, A7: 50, B7: "b", D7: 400 },
    expected: 50,
    source: "Excel",
    note: 'MAXIFS function, example 3: the largest weight (A) among grade "b" (B) rows whose level (D) exceeds 100 — rows 5 and 7, weights 1 and 50; Microsoft prints 50.',
  },
  {
    formula: 'MAXIFS(A2:A6,B2:B6,"a",D2:D6,">200")',
    cells: { A2: 10, B2: "b", D2: 100, A3: 1, B3: "a", D3: 100, A4: 100, B4: "a", D4: 200, A5: 1, B5: "b", D5: 300, A6: 1, B6: "a", D6: 100 },
    expected: 0,
    source: "Excel",
    note: 'MAXIFS function, example 3 variant: grade "a" rows (3, 4, 6) have levels 100, 200 and 100 — none exceeds 200, so nothing matches; Microsoft prints 0.',
  },
  {
    formula: 'MAXIFS(A1:A2,B1:B3,"a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'MAXIFS function Remarks: "The size and shape of the max_range and criteria_rangeN arguments must be the same, otherwise these functions return the #VALUE! error"; A1:A2 (2 rows) and B1:B3 (3 rows) differ.',
  },

  // --- MIN: number1, [number2], … --------------------------------------
  {
    formula: "MIN(A2:A6)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 },
    expected: 2,
    source: "Excel",
    note: "MIN function, example: the smallest of 10, 7, 9, 27, 2 in A2:A6; Microsoft prints 2.",
  },
  {
    formula: "MIN(A2:A6,0)",
    cells: { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 },
    expected: 0,
    source: "Excel",
    note: "MIN function, example: the same five numbers plus the literal 0, which becomes the smallest; Microsoft prints 0.",
  },
  {
    formula: 'MIN("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'MIN function Remarks: "Arguments that are error values or text that cannot be translated into numbers cause errors"; a direct, non-numeric text argument is #VALUE!.',
  },

  // --- MINA: value1, [value2], … -----------------------------------------
  {
    formula: "MINA(A2:A6)",
    cells: { A2: false, A3: 0.2, A4: 0.5, A5: 0.4, A6: 0.8 },
    expected: 0,
    source: "Excel",
    note: "MINA function, example: the smallest of FALSE (counted as 0), 0.2, 0.5, 0.4 and 0.8 in A2:A6; Microsoft prints 0.",
  },
  {
    formula: 'MINA("a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'MINA function Remarks: "Arguments that are error values or text that cannot be translated into numbers cause errors"; a direct, non-numeric text argument is #VALUE!.',
  },
  {
    formula: "MINA(TRUE,0.3,FALSE)",
    expected: 0,
    source: "pinned",
    note: "pinned: MINA counts TRUE as 1 and FALSE as 0, so the smallest of 1, 0.3 and 0 is 0 exactly.",
  },

  // --- MINIFS: min_range, criteria_range1, criteria1, … -------------------
  {
    formula: "MINIFS(A2:A7,B2:B7,1)",
    cells: { A2: 89, B2: 1, A3: 93, B3: 2, A4: 96, B4: 2, A5: 85, B5: 3, A6: 91, B6: 1, A7: 88, B7: 1 },
    expected: 88,
    source: "Excel",
    note: "MINIFS function, example: the smallest grade (A) among the rows weighted 1 (B) — 89, 91 and 88; Microsoft prints 88.",
  },
  {
    formula: 'MINIFS(A2:A7,B2:B7,"b",D2:D7,">100")',
    cells: { A2: 10, B2: "b", D2: 100, A3: 11, B3: "a", D3: 100, A4: 12, B4: "a", D4: 200, A5: 13, B5: "b", D5: 300, A6: 14, B6: "b", D6: 300, A7: 15, B7: "b", D7: 400 },
    expected: 13,
    source: "Excel",
    note: 'MINIFS function, example: the smallest weight (A) among grade "b" (B) rows whose level (D) exceeds 100 — rows 5, 6 and 7, weights 13, 14 and 15; Microsoft prints 13.',
  },
  {
    formula: 'MINIFS(A1:A2,B1:B3,"a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "MINIFS function Remarks: the size and shape of the min_range and every criteria_range argument must match, otherwise #VALUE!; A1:A2 (2 rows) and B1:B3 (3 rows) differ.",
  },

  // The STDEV*/VAR* pages share one ten-value "breaking strength" data set
  // (1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299), but not one
  // address: STDEV, STDEV.P, STDEVP, STDEVA and STDEVPA's tables carry a
  // "Data" title row above the "Strength" header, so the values sit at
  // A3:A12; the other pages have the header only, A2:A11. Each checked on
  // its own page. Pinned rows use data with exact results (population
  // variance of 2,4,4,4,5,5,7,9 is 4, sample variance of 1…5 is 2.5).
  // --- STDEV: number1, [number2], ... (legacy, own page) --------------------
  {
    formula: "STDEV(A3:A12)",
    cells: { A3: 1345, A4: 1301, A5: 1368, A6: 1322, A7: 1310, A8: 1370, A9: 1318, A10: 1350, A11: 1303, A12: 1299 },
    expected: 27.46392,
    digits: 7,
    source: "Excel",
    note: "STDEV function, example: sample standard deviation of 10 tools' breaking strength, data at A3:A12 (this page's table has a \"Data\" title row above the \"Strength\" header); Microsoft prints 27.46392.",
  },
  {
    formula: "STDEV(1,2,3)",
    expected: 1,
    source: "pinned",
    note: "pinned: sample variance of 1,2,3 is ((1−2)²+(2−2)²+(3−2)²)/(3−1) = 2/2 = 1 (mean 2), so STDEV(1,2,3) = √1 = 1 exactly.",
  },
  {
    formula: "STDEV(5)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the sample form divides the sum of squared deviations by n−1 (the page's own \"n-1 method\" remark); a single value leaves n−1=0, so STDEV(5) is #DIV/0!.",
  },

  // --- STDEV.P: number1, [number2], ... --------------------------------------
  {
    formula: "STDEV.P(A3:A12)",
    cells: { A3: 1345, A4: 1301, A5: 1368, A6: 1322, A7: 1310, A8: 1370, A9: 1318, A10: 1350, A11: 1303, A12: 1299 },
    expected: 26.05455814,
    digits: 10,
    source: "Excel",
    note: "STDEV.P function, example: population standard deviation of the same breaking-strength data, data at A3:A12 (this page keeps the \"Data\"/\"Strength\" double header too); Microsoft prints 26.05455814.",
  },
  {
    formula: "STDEV.P(2,4,4,4,5,5,7,9)",
    expected: 2,
    source: "pinned",
    note: "pinned: population variance of 2,4,4,4,5,5,7,9 is 32/8 = 4 (mean 5; squared deviations 9,1,1,1,0,0,4,16 sum to 32), so STDEV.P(…) = √4 = 2 exactly.",
  },
  {
    formula: "STDEV.P(A1:A3)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the population form needs at least one number (it divides by n); A1:A3 holds no values at all, so STDEV.P(A1:A3) is #DIV/0!.",
  },

  // --- STDEV.S: number1, [number2], ... --------------------------------------
  {
    formula: "STDEV.S(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 27.46391572,
    digits: 10,
    source: "Excel",
    note: "STDEV.S function, example: sample standard deviation of the breaking-strength data, data at A2:A11 (single \"Strength\" header, one row higher than STDEV's own page); Microsoft prints 27.46391572.",
  },
  {
    formula: "STDEV.S(1,2,3)",
    expected: 1,
    source: "pinned",
    note: "pinned: sample variance of 1,2,3 is ((1−2)²+(2−2)²+(3−2)²)/(3−1) = 2/2 = 1 (mean 2), so STDEV.S(1,2,3) = √1 = 1 exactly.",
  },
  {
    formula: "STDEV.S(5)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the sample form divides by n−1; a single value leaves n−1=0, so STDEV.S(5) is #DIV/0!.",
  },

  // --- STDEVA: value1, [value2], ... ------------------------------------------
  {
    formula: "STDEVA(A3:A12)",
    cells: { A3: 1345, A4: 1301, A5: 1368, A6: 1322, A7: 1310, A8: 1370, A9: 1318, A10: 1350, A11: 1303, A12: 1299 },
    expected: 27.46391572,
    digits: 10,
    source: "Excel",
    note: "STDEVA function, example: the same breaking-strength data (all plain numbers, so STDEVA behaves like STDEV.S), data at A3:A12 (double header, like STDEV's own page); Microsoft prints 27.46391572.",
  },
  {
    formula: "STDEVA(1,5,TRUE,TRUE)",
    expected: 2,
    source: "pinned",
    note:
      "pinned: STDEVA counts a directly-typed TRUE as 1 (Remarks: \"Logical values … that you type directly into the list of arguments are counted\"); reading 1,5,TRUE,TRUE as 1,5,1,1, the sample variance is " +
      "((1−2)²+(5−2)²+(1−2)²+(1−2)²)/(4−1) = 12/3 = 4 (mean 2), so STDEVA(1,5,TRUE,TRUE) = √4 = 2 exactly.",
  },
  {
    formula: "STDEVA(5)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the sample form divides by n−1; a single value leaves n−1=0, so STDEVA(5) is #DIV/0!.",
  },

  // --- STDEVP: number1, [number2], ... (legacy, own page) --------------------
  {
    formula: "STDEVP(A3:A12)",
    cells: { A3: 1345, A4: 1301, A5: 1368, A6: 1322, A7: 1310, A8: 1370, A9: 1318, A10: 1350, A11: 1303, A12: 1299 },
    expected: 26.05456,
    digits: 7,
    source: "Excel",
    note: "STDEVP function, example: population standard deviation of the breaking-strength data, data at A3:A12 (\"Data\"/\"Strength\" double header); Microsoft prints 26.05456.",
  },
  {
    formula: "STDEVP(2,4,4,4,5,5,7,9)",
    expected: 2,
    source: "pinned",
    note: "pinned: population variance of 2,4,4,4,5,5,7,9 is 32/8 = 4 (mean 5; squared deviations sum to 32), so STDEVP(…) = √4 = 2 exactly.",
  },
  {
    formula: "STDEVP(A1:A3)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the population form needs at least one number (it divides by n); A1:A3 holds no values at all, so STDEVP(A1:A3) is #DIV/0!.",
  },

  // --- STDEVPA: value1, [value2], ... -----------------------------------------
  {
    formula: "STDEVPA(A3:A12)",
    cells: { A3: 1345, A4: 1301, A5: 1368, A6: 1322, A7: 1310, A8: 1370, A9: 1318, A10: 1350, A11: 1303, A12: 1299 },
    expected: 26.05455814,
    digits: 10,
    source: "Excel",
    note: "STDEVPA function, example: the same breaking-strength data (all plain numbers), data at A3:A12; the Description states the more precise 26.05455814 where the Result cell shows the rounded 26.05456.",
  },
  {
    formula: "STDEVPA(FALSE,FALSE,4,4)",
    expected: 2,
    source: "pinned",
    note:
      "pinned: STDEVPA counts a directly-typed FALSE as 0; reading FALSE,FALSE,4,4 as 0,0,4,4, the population variance is " +
      "((0−2)²+(0−2)²+(4−2)²+(4−2)²)/4 = 16/4 = 4 (mean 2), so STDEVPA(FALSE,FALSE,4,4) = √4 = 2 exactly.",
  },
  {
    formula: "STDEVPA(A1:A3)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the population form needs at least one number; an all-empty A1:A3 (empty cells are skipped even by the -A reading) holds none, so STDEVPA(A1:A3) is #DIV/0!.",
  },

  // --- VAR: number1, [number2], ... (legacy, own page) ------------------------
  {
    formula: "VAR(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 754.2667,
    digits: 7,
    source: "Excel",
    note: "VAR function, example: sample variance of the breaking-strength data, data at A2:A11 (single \"Strength\" header); Microsoft prints 754.2667.",
  },
  {
    formula: "VAR(1,2,3,4,5)",
    expected: 2.5,
    source: "pinned",
    note: "pinned: sample variance of 1,2,3,4,5 (mean 3) is ((−2)²+(−1)²+0²+1²+2²)/(5−1) = 10/4 = 2.5 exactly.",
  },
  {
    formula: "VAR(5)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the sample form divides by n−1; a single value leaves n−1=0, so VAR(5) is #DIV/0!.",
  },

  // --- VAR.P: number1, [number2], ... ------------------------------------------
  {
    formula: "VAR.P(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 678.84,
    digits: 5,
    source: "Excel",
    note: "VAR.P function, example: population variance of the breaking-strength data, data at A2:A11; Microsoft prints 678.84.",
  },
  {
    formula: "VAR.P(2,4,4,4,5,5,7,9)",
    expected: 4,
    source: "pinned",
    note: "pinned: population variance of 2,4,4,4,5,5,7,9 (mean 5) is (9+1+1+1+0+0+4+16)/8 = 32/8 = 4 exactly.",
  },
  {
    formula: "VAR.P(A1:A3)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the population form needs at least one number (it divides by n); A1:A3 holds no values at all, so VAR.P(A1:A3) is #DIV/0!.",
  },

  // --- VAR.S: number1, [number2], ... ------------------------------------------
  {
    formula: "VAR.S(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 754.27,
    digits: 5,
    source: "Excel",
    note: "VAR.S function, example: sample variance of the breaking-strength data, data at A2:A11; Microsoft prints 754.27 (the same page's VAR.P(A2:A11) row is transcribed under VAR.P).",
  },
  {
    formula: "VAR.S(1,2,3,4,5)",
    expected: 2.5,
    source: "pinned",
    note: "pinned: sample variance of 1,2,3,4,5 (mean 3) is ((−2)²+(−1)²+0²+1²+2²)/(5−1) = 10/4 = 2.5 exactly.",
  },
  {
    formula: "VAR.S(5)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the sample form divides by n−1; a single value leaves n−1=0, so VAR.S(5) is #DIV/0!.",
  },

  // --- VARA: value1, [value2], ... ----------------------------------------------
  {
    formula: "VARA(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 754.26667,
    digits: 8,
    source: "Excel",
    note: "VARA function, example: the same breaking-strength data (all plain numbers, so VARA behaves like VAR.S), data at A2:A11; Microsoft prints 754.26667.",
  },
  {
    formula: "VARA(1,5,TRUE,TRUE)",
    expected: 4,
    source: "pinned",
    note:
      "pinned: VARA counts a directly-typed TRUE as 1; reading 1,5,TRUE,TRUE as 1,5,1,1 (mean 2), the sample variance is " +
      "((1−2)²+(5−2)²+(1−2)²+(1−2)²)/(4−1) = 12/3 = 4 exactly.",
  },
  {
    formula: "VARA(5)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the sample form divides by n−1; a single value leaves n−1=0, so VARA(5) is #DIV/0!.",
  },

  // --- VARP: number1, [number2], ... (legacy, own page) --------------------------
  {
    formula: "VARP(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 678.84,
    digits: 5,
    source: "Excel",
    note: "VARP function, example: population variance of the breaking-strength data, data at A2:A11; Microsoft prints 678.84.",
  },
  {
    formula: "VARP(2,4,4,4,5,5,7,9)",
    expected: 4,
    source: "pinned",
    note: "pinned: population variance of 2,4,4,4,5,5,7,9 (mean 5) is (9+1+1+1+0+0+4+16)/8 = 32/8 = 4 exactly.",
  },
  {
    formula: "VARP(A1:A3)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the population form needs at least one number; A1:A3 holds no values at all, so VARP(A1:A3) is #DIV/0!.",
  },

  // --- VARPA: value1, [value2], ... -----------------------------------------------
  {
    formula: "VARPA(A2:A11)",
    cells: { A2: 1345, A3: 1301, A4: 1368, A5: 1322, A6: 1310, A7: 1370, A8: 1318, A9: 1350, A10: 1303, A11: 1299 },
    expected: 678.84,
    digits: 5,
    source: "Excel",
    note: "VARPA function, example: the same breaking-strength data (all plain numbers), data at A2:A11; Microsoft prints 678.84.",
  },
  {
    formula: "VARPA(FALSE,FALSE,4,4)",
    expected: 4,
    source: "pinned",
    note:
      "pinned: VARPA counts a directly-typed FALSE as 0; reading FALSE,FALSE,4,4 as 0,0,4,4 (mean 2), the population variance is " +
      "((0−2)²+(0−2)²+(4−2)²+(4−2)²)/4 = 16/4 = 4 exactly.",
  },
  {
    formula: "VARPA(A1:A3)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: the population form needs at least one number; an all-empty A1:A3 (empty cells are skipped even by the -A reading) holds none, so VARPA(A1:A3) is #DIV/0!.",
  },

  // --- AVEDEV: number1, [number2], ... ---------------------------------------------
  {
    formula: "AVEDEV(A2:A8)",
    cells: { A2: 4, A3: 5, A4: 6, A5: 7, A6: 5, A7: 4, A8: 3 },
    expected: 1.020408,
    digits: 7,
    source: "Excel",
    note: "AVEDEV function, example: mean absolute deviation of 4,5,6,7,5,4,3 from their mean; Microsoft prints 1.020408.",
  },
  {
    formula: "AVEDEV(2,4,4,4,5,5,7,9)",
    expected: 1.5,
    source: "pinned",
    note: "pinned: mean of 2,4,4,4,5,5,7,9 is 5; absolute deviations 3,1,1,1,0,0,2,4 sum to 12, so AVEDEV(…) = 12/8 = 1.5 exactly.",
  },
  {
    formula: "AVEDEV()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: a call outside the function's argument count is #VALUE! (evaluate.ts; Excel refuses to enter it); AVEDEV needs at least one argument, so AVEDEV() is #VALUE! (its page's Remarks name no error value).",
  },

  // --- DEVSQ: number1, [number2], ... ----------------------------------------------
  {
    formula: "DEVSQ(A2:A8)",
    cells: { A2: 4, A3: 5, A4: 8, A5: 7, A6: 11, A7: 4, A8: 3 },
    expected: 48,
    source: "Excel",
    note: "DEVSQ function, example: sum of squared deviations of 4,5,8,7,11,4,3 from their mean; Microsoft prints 48.",
  },
  {
    formula: "DEVSQ(1,2,3)",
    expected: 2,
    source: "pinned",
    note: "pinned: mean of 1,2,3 is 2; squared deviations 1,0,1 sum to 2, so DEVSQ(1,2,3) = 2 exactly (DEVSQ never divides, so any integer input is exact).",
  },
  {
    formula: 'DEVSQ(1,"abc")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: "DEVSQ Remarks, verbatim: \"Arguments that are error values or text that cannot be translated into numbers cause errors.\"; \"abc\" typed directly cannot become a number.",
  },

  // --- GEOMEAN: number1, [number2], ... --------------------------------------------
  {
    formula: "GEOMEAN(A2:A8)",
    cells: { A2: 4, A3: 5, A4: 8, A5: 7, A6: 11, A7: 4, A8: 3 },
    expected: 5.476987,
    digits: 7,
    source: "Excel",
    note: "GEOMEAN function, example: geometric mean of 4,5,8,7,11,4,3; Microsoft prints 5.476987.",
  },
  {
    formula: "GEOMEAN(4,9)",
    expected: 6,
    source: "pinned",
    note: "pinned: the geometric mean of 4 and 9 is √(4·9) = √36 = 6 exactly.",
  },
  {
    formula: "GEOMEAN(1,-1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "GEOMEAN Remarks, verbatim: \"If any data point ≤ 0, GEOMEAN returns the #NUM! error value.\" (−1 ≤ 0 here).",
  },

  // --- HARMEAN: number1, [number2], ... --------------------------------------------
  {
    formula: "HARMEAN(A2:A8)",
    cells: { A2: 4, A3: 5, A4: 8, A5: 7, A6: 11, A7: 4, A8: 3 },
    expected: 5.028376,
    digits: 7,
    source: "Excel",
    note: "HARMEAN function, example: harmonic mean of 4,5,8,7,11,4,3; Microsoft prints 5.028376.",
  },
  {
    formula: "HARMEAN(1,3)",
    expected: 1.5,
    source: "pinned",
    note: "pinned: the harmonic mean of 1 and 3 is 2/(1/1+1/3) = 2/(4/3) = 1.5 exactly.",
  },
  {
    formula: "HARMEAN(1,-1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: "HARMEAN Remarks, verbatim: \"If any data point ≤ 0, HARMEAN returns the #NUM! error value.\" (−1 ≤ 0 here).",
  },

  // PERCENTILE, PERCENTILE.EXC and FREQUENCY print their examples only as
  // screenshots, so their Excel rows (where they have any) come from the
  // Remarks' error conditions; FREQUENCY's Remarks name none, so it is in
  // conformance.spec.ts's list of pages without an example table.
  // --- MEDIAN: number1, [number2], ... -----------------------------------
  {
    formula: "MEDIAN(A2:A6)",
    cells: { A2: 1, A3: 2, A4: 3, A5: 4, A6: 5 },
    expected: 3,
    source: "Excel",
    note: "MEDIAN function, example row 1: median of the 5 numbers 1-5; the third value (3) is the median.",
  },
  {
    formula: "MEDIAN(A2:A7)",
    cells: { A2: 1, A3: 2, A4: 3, A5: 4, A6: 5, A7: 6 },
    expected: 3.5,
    source: "Excel",
    note: "MEDIAN function, example row 2: median of the 6 numbers 1-6; an even count, so the midpoint between the 3rd and 4th values (3 and 4).",
  },
  {
    formula: "MEDIAN()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (0, MEDIAN needs at least number1); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // --- MODE: number1, [number2], ... (legacy, shares MODE.SNGL's rule) ---
  {
    formula: "MODE(A2:A7)",
    cells: { A2: 5.6, A3: 4, A4: 4, A5: 3, A6: 2, A7: 4 },
    expected: 4,
    source: "Excel",
    note: "MODE function, example: of 5.6, 4, 4, 3, 2, 4, the value 4 repeats three times and is the mode.",
  },
  {
    formula: "MODE(1,2,3)",
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'MODE Remarks, verbatim: "If the data set contains no duplicate data points, MODE returns the #N/A error value." — none of 1, 2, 3 repeats.',
  },
  {
    formula: "MODE(7,7,7,2,9)",
    expected: 7,
    source: "pinned",
    note: "pinned: 7 occurs three times against one each for 2 and 9, so it is unambiguously the most frequent value.",
  },

  // --- MODE.SNGL: number1, [number2], ... --------------------------------
  {
    formula: "MODE.SNGL(A2:A7)",
    cells: { A2: 5.6, A3: 4, A4: 4, A5: 3, A6: 2, A7: 4 },
    expected: 4,
    source: "Excel",
    note: "MODE.SNGL function, example: the same data as MODE's page; 4 repeats three times and is the mode.",
  },
  {
    formula: "MODE.SNGL(1,2,3)",
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'MODE.SNGL Remarks, verbatim: "If the data set contains no duplicate data points, MODE.SNGL returns the #N/A error value." — none of 1, 2, 3 repeats.',
  },
  {
    formula: "MODE.SNGL(5,5,5,1,2)",
    expected: 5,
    source: "pinned",
    note: "pinned: 5 occurs three times against one each for 1 and 2, so it is unambiguously the most frequent value.",
  },

  // --- LARGE: array, k -----------------------------------------------------
  {
    formula: "LARGE(A2:B6,3)",
    cells: { A2: 3, B2: 4, A3: 5, B3: 2, A4: 3, B4: 4, A5: 5, B5: 6, A6: 4, B6: 7 },
    expected: 5,
    source: "Excel",
    note: "LARGE function, example row 1: 3rd largest of the ten numbers in A2:B6 (sorted descending 7,6,5,5,4,4,4,3,3,2).",
  },
  {
    formula: "LARGE(A2:B6,7)",
    cells: { A2: 3, B2: 4, A3: 5, B3: 2, A4: 3, B4: 4, A5: 5, B5: 6, A6: 4, B6: 7 },
    expected: 4,
    source: "Excel",
    note: "LARGE function, example row 2: 7th largest of the same ten numbers.",
  },
  {
    formula: "LARGE({1,2,3},4)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'LARGE Remarks, verbatim: "If k ≤ 0 or if k is greater than the number of data points, LARGE returns the #NUM! error value." — k=4 exceeds the 3 data points.',
  },

  // --- SMALL: array, k -------------------------------------------------------
  {
    formula: "SMALL(A2:A10,4)",
    cells: { A2: 3, B2: 1, A3: 4, B3: 4, A4: 5, B4: 8, A5: 2, B5: 3, A6: 3, B6: 7, A7: 4, B7: 12, A8: 6, B8: 54, A9: 4, B9: 8, A10: 7, B10: 23 },
    expected: 4,
    source: "Excel",
    note: "SMALL function, example row 1: 4th smallest of A2:A10 (sorted ascending 2,3,3,4,4,4,5,6,7).",
  },
  {
    formula: "SMALL(B2:B10,2)",
    cells: { A2: 3, B2: 1, A3: 4, B3: 4, A4: 5, B4: 8, A5: 2, B5: 3, A6: 3, B6: 7, A7: 4, B7: 12, A8: 6, B8: 54, A9: 4, B9: 8, A10: 7, B10: 23 },
    expected: 3,
    source: "Excel",
    note: "SMALL function, example row 2: 2nd smallest of B2:B10 (sorted ascending 1,3,4,7,8,8,12,23,54).",
  },
  {
    formula: "SMALL({1,2,3},4)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'SMALL Remarks, verbatim: "If k ≤ 0 or if k exceeds the number of data points, SMALL returns the #NUM! error value." — k=4 exceeds the 3 data points.',
  },

  // --- PERCENTILE: array, k (legacy; page's example is an image only) -------
  {
    formula: "PERCENTILE({1,2,3,4},0.5)",
    expected: 2.5,
    source: "pinned",
    note: "pinned: PERCENTILE's page example is a screenshot with no printed numbers; by the stated rule (k·(n−1) interpolated), k=0.5 over {1,2,3,4} ranks at 0.5·3=1.5, halfway between 2 and 3.",
  },
  {
    formula: "PERCENTILE({1,2,3},1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERCENTILE Remarks, verbatim: "If k is < 0 or if k > 1, PERCENTILE returns the #NUM! error value." — k=1.5.',
  },
  {
    formula: 'PERCENTILE({1,2,3},"x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'PERCENTILE Remarks, verbatim: "If k is non-numeric, PERCENTILE returns the #VALUE! error value." — k is the text "x".',
  },

  // --- PERCENTILE.EXC: array, k (its example is an image only) --------------
  {
    formula: "PERCENTILE.EXC({1,2,3,4},0.5)",
    expected: 2.5,
    source: "pinned",
    note: "pinned: PERCENTILE.EXC's page example is a screenshot with no printed numbers; by the stated rule (k·(n+1) interpolated), k=0.5 over {1,2,3,4} ranks at 0.5·5=2.5, halfway between 2 and 3.",
  },
  {
    formula: "PERCENTILE.EXC({1,2,3},1)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERCENTILE.EXC Remarks, verbatim: "If k ≤ 0 or if k ≥ 1, PERCENTILE.EXC returns the #NUM! error value." — k=1.',
  },
  {
    formula: 'PERCENTILE.EXC({1,2,3},"x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'PERCENTILE.EXC Remarks, verbatim: "If k is nonnumeric, PERCENTILE.EXC returns the #VALUE! error value." — k is the text "x".',
  },

  // --- PERCENTILE.INC: array, k ----------------------------------------------
  {
    formula: "PERCENTILE.INC(A2:A5,0.3)",
    cells: { A2: 1, A3: 2, A4: 3, A5: 4 },
    expected: 1.9,
    source: "Excel",
    note: "PERCENTILE.INC function, example: 30th percentile of 1,2,3,4; Microsoft prints 1.9.",
  },
  {
    formula: "PERCENTILE.INC({1,2,3},1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERCENTILE.INC Remarks, verbatim: "If k is < 0 or if k > 1, PERCENTILE.INC returns the #NUM! error value." — k=1.5.',
  },
  {
    formula: "PERCENTILE.INC({2,4,6,8},0.5)",
    expected: 5,
    source: "pinned",
    note: "pinned: k=0.5 over {2,4,6,8} ranks at 0.5·3=1.5, halfway between 4 and 6, giving the exact median 5.",
  },

  // --- QUARTILE: array, quart (legacy) ---------------------------------------
  {
    formula: "QUARTILE(A2:A9,1)",
    cells: { A2: 1, A3: 2, A4: 4, A5: 7, A6: 8, A7: 9, A8: 10, A9: 12 },
    expected: 3.5,
    source: "Excel",
    note: "QUARTILE function, example: first quartile (25th percentile) of 1,2,4,7,8,9,10,12; Microsoft prints 3.5.",
  },
  {
    formula: "QUARTILE({1,2,3},5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'QUARTILE Remarks, verbatim: "If quart < 0 or if quart > 4, QUARTILE returns the #NUM! error value." — quart=5.',
  },
  {
    formula: "QUARTILE({1,2,3,4},2)",
    expected: 2.5,
    source: "pinned",
    note: "pinned: quart=2 is the median; QUARTILE(array,2) is PERCENTILE.INC(array,0.5), which over {1,2,3,4} ranks at 1.5, halfway between 2 and 3.",
  },

  // --- QUARTILE.EXC: array, quart ---------------------------------------------
  {
    formula: "QUARTILE.EXC(A2:A12,1)",
    cells: { A2: 6, A3: 7, A4: 15, A5: 36, A6: 39, A7: 40, A8: 41, A9: 42, A10: 43, A11: 47, A12: 49 },
    expected: 15,
    source: "Excel",
    note: "QUARTILE.EXC function, example row 1: first quartile of the 11 values in A2:A12; Microsoft prints 15 (rank 0.25·12=3, the 3rd smallest).",
  },
  {
    formula: "QUARTILE.EXC(A2:A12,3)",
    cells: { A2: 6, A3: 7, A4: 15, A5: 36, A6: 39, A7: 40, A8: 41, A9: 42, A10: 43, A11: 47, A12: 49 },
    expected: 43,
    source: "Excel",
    note: "QUARTILE.EXC function, example row 2: third quartile of the same 11 values; Microsoft prints 43 (rank 0.75·12=9, the 9th smallest).",
  },
  {
    formula: "QUARTILE.EXC({1,2,3},4)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'QUARTILE.EXC Remarks, verbatim: "If quart ≤ 0 or if quart ≥ 4, QUARTILE.EXC returns the #NUM! error value." — quart=4.',
  },

  // --- QUARTILE.INC: array, quart ---------------------------------------------
  {
    formula: "QUARTILE.INC(A2:A9,1)",
    cells: { A2: 1, A3: 2, A4: 4, A5: 7, A6: 8, A7: 9, A8: 10, A9: 12 },
    expected: 3.5,
    source: "Excel",
    note: "QUARTILE.INC function, example: first quartile of 1,2,4,7,8,9,10,12, the same data as QUARTILE's page; Microsoft prints 3.5.",
  },
  {
    formula: "QUARTILE.INC({1,2,3},5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'QUARTILE.INC Remarks, verbatim: "If quart < 0 or if quart > 4, QUARTILE.INC returns the #NUM! error value." — quart=5.',
  },
  {
    formula: "QUARTILE.INC({2,4,6,8},2)",
    expected: 5,
    source: "pinned",
    note: "pinned: quart=2 is the median; over {2,4,6,8} that is PERCENTILE.INC-equivalent at rank 1.5, halfway between 4 and 6, giving 5 exactly.",
  },

  // --- PERCENTRANK: array, x, [significance] (legacy) -------------------------
  {
    formula: "PERCENTRANK(A2:A11,2)",
    cells: { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 },
    expected: 0.333,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK function, example row 1: rank of 2 among the ten values; three values (1,1,1) are below it, 3/9=0.333.",
  },
  {
    formula: "PERCENTRANK(A2:A11,5)",
    cells: { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 },
    expected: 0.583,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK function, example row 4: 5 is absent from the data, so the rank is interpolated one-quarter of the way between the ranks of 4 and 8; Microsoft prints 0.583.",
  },
  {
    formula: "PERCENTRANK(A2:A11,8)",
    cells: { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 },
    expected: 0.666,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK function, example row 3: rank of 8; six values are below it, 6/9=0.666.",
  },
  {
    formula: "PERCENTRANK({1,2,3},2,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERCENTRANK Remarks, verbatim: "If significance < 1, PERCENTRANK returns the #NUM! error value." — significance=0.',
  },

  // --- PERCENTRANK.EXC: array, x, [significance] --------------------------------
  {
    formula: "PERCENTRANK.EXC(A2:A10,7)",
    cells: { A2: 1, A3: 2, A4: 3, A5: 6, A6: 6, A7: 6, A8: 7, A9: 8, A10: 9 },
    expected: 0.7,
    source: "Excel",
    note: "PERCENTRANK.EXC function, example row 1: rank of 7 among the nine values 1,2,3,6,6,6,7,8,9; (6 below + 1)/(9+1)=0.7.",
  },
  {
    formula: "PERCENTRANK.EXC(A2:A10,5.43)",
    cells: { A2: 1, A3: 2, A4: 3, A5: 6, A6: 6, A7: 6, A8: 7, A9: 8, A10: 9 },
    expected: 0.381,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK.EXC function, example row 2: 5.43 is absent, interpolated 81% of the way from 3's rank (0.3) to 6's rank (0.4); Microsoft prints 0.381.",
  },
  {
    formula: "PERCENTRANK.EXC(A2:A10,5.43,1)",
    cells: { A2: 1, A3: 2, A4: 3, A5: 6, A6: 6, A7: 6, A8: 7, A9: 8, A10: 9 },
    expected: 0.3,
    digits: 1,
    source: "Excel",
    note: "PERCENTRANK.EXC function, example row 3: the same rank of 5.43, truncated to 1 significant digit instead of the default 3; Microsoft prints 0.3.",
  },
  {
    formula: "PERCENTRANK.EXC({1,2,3},2,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERCENTRANK.EXC Remarks, verbatim: "If significance < 1, PERCENTRANK.EXC returns the #NUM! error value." — significance=0.',
  },

  // --- PERCENTRANK.INC: array, x, [significance] --------------------------------
  {
    formula: "PERCENTRANK.INC(A2:A11,2)",
    cells: { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 },
    expected: 0.333,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK.INC function, example: rank of 2 (3 values below it of 10, 3/9=0.333), the same data as PERCENTRANK's page.",
  },
  {
    formula: "PERCENTRANK.INC(A2:A11,5)",
    cells: { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 },
    expected: 0.583,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK.INC function, example: 5 is absent, interpolated one-quarter of the way between the ranks of 4 and 8; Microsoft prints 0.583.",
  },
  {
    formula: "PERCENTRANK.INC(A2:A11,8)",
    cells: { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 },
    expected: 0.666,
    digits: 3,
    source: "Excel",
    note: "PERCENTRANK.INC function, example: rank of 8 (6 values below it of 10, 6/9=0.666).",
  },
  {
    formula: "PERCENTRANK.INC({1,2,3},2,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'PERCENTRANK.INC Remarks, verbatim: "If significance < 1, PERCENTRANK.INC returns the #NUM! error value." — significance=0.',
  },

  // --- RANK: number, ref, [order] (legacy; ref must be a reference) -----------
  {
    formula: "RANK(A3,A2:A6,1)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: 3,
    source: "Excel",
    note: "RANK function, example row 1: rank of 3.5 among 7,3.5,3.5,1,2 sorted ascending (order=1); Microsoft prints 3.",
  },
  {
    formula: "RANK(A2,A2:A6,1)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: 5,
    source: "Excel",
    note: "RANK function, example row 2: rank of 7 (the largest) in the same ascending list; Microsoft prints 5.",
  },
  {
    formula: "RANK(100,A2:A6)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: 100 does not occur in A2:A6; SPEC-127's Function pins — RANK is #N/A when the number is absent from ref (evaluate.ts rankImpl).",
  },

  // --- RANK.AVG: number, ref, [order] ---------------------------------------------
  {
    formula: "RANK.AVG(94,B2:B8)",
    cells: { B2: 89, B3: 88, B4: 92, B5: 101, B6: 94, B7: 97, B8: 95 },
    expected: 4,
    source: "Excel",
    note: "RANK.AVG function, example: rank of the temperature 94 among B2:B8 (89,88,92,101,94,97,95), descending by default; Microsoft prints 4 (the 4th hottest day).",
  },
  {
    formula: "RANK.AVG(1000,B2:B8)",
    cells: { B2: 89, B3: 88, B4: 92, B5: 101, B6: 94, B7: 97, B8: 95 },
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: 1000 does not occur in B2:B8; SPEC-127's Function pins — RANK.AVG is #N/A when the number is absent from ref.",
  },
  {
    formula: "RANK.AVG(C1,C1:C4)",
    cells: { C1: 5, C2: 5, C3: 3, C4: 1 },
    expected: 1.5,
    source: "pinned",
    note: "pinned: ranking 5 descending among 5,5,3,1 — nothing is larger (before=0), two values tie at 5 (equal=2), so RANK.AVG averages ranks 1 and 2: 0+(2+1)/2=1.5.",
  },

  // --- RANK.EQ: number, ref, [order] ------------------------------------------------
  {
    formula: "RANK.EQ(A2,A2:A6,1)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: 5,
    source: "Excel",
    note: "RANK.EQ function, example row 1: rank of 7 among 7,3.5,3.5,1,2 sorted ascending (order=1, non-zero); Microsoft prints 5.",
  },
  {
    formula: "RANK.EQ(A6,A2:A6)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: 4,
    source: "Excel",
    note: "RANK.EQ function, example row 2: rank of 2, order omitted so descending by default; Microsoft prints 4.",
  },
  {
    formula: "RANK.EQ(A3,A2:A6,1)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: 3,
    source: "Excel",
    note: "RANK.EQ function, example row 3: rank of 3.5 (tied with A4) ascending; Microsoft prints 3, the top rank of the tied pair.",
  },
  {
    formula: "RANK.EQ(100,A2:A6)",
    cells: { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 },
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: 100 does not occur in A2:A6; SPEC-127's Function pins — RANK.EQ is #N/A when the number is absent from ref.",
  },

  // --- TRIMMEAN: array, percent ------------------------------------------------------
  {
    formula: "TRIMMEAN(A2:A12,0.2)",
    cells: { A2: 4, A3: 5, A4: 6, A5: 7, A6: 2, A7: 3, A8: 4, A9: 5, A10: 1, A11: 2, A12: 3 },
    expected: 3.778,
    digits: 4,
    source: "Excel",
    note: "TRIMMEAN function, example: mean of A2:A12 (11 values) with 20% trimmed (1 value off each end, leaving 2,2,3,3,4,4,5,5,6); Microsoft prints 3.778.",
  },
  {
    formula: "TRIMMEAN({1,2,3},1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'TRIMMEAN Remarks, verbatim: "If percent < 0 or percent > 1, TRIMMEAN returns the #NUM! error value." — percent=1.5.',
  },
  {
    formula: "TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.2)",
    expected: 5.5,
    source: "pinned",
    note: "pinned: 10 values, 20% trims floor(10·0.2/2)=1 from each end, leaving 2..9 (sum 44 over 8 values) for a mean of exactly 5.5.",
  },

  // --- FREQUENCY: data_array, bins_array (page's example is an image only) ----
  {
    formula: "FREQUENCY(A1:A10,C1:C2)",
    cells: { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, A6: 6, A7: 7, A8: 8, A9: 9, A10: 10, C1: 3, C2: 6 },
    array: "D1:D3",
    expected: 3,
    expectedArray: [[3], [3], [4]],
    source: "pinned",
    note: "pinned: FREQUENCY's page shows only a screenshot with no printed numbers; by the stated rule, bins 3 and 6 over 1..10 count {1,2,3}=3, {4,5,6}=3, and {7,8,9,10}=4 above the last bin.",
  },
  {
    formula: "FREQUENCY(A1:A10,C1:C2)",
    cells: { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, A6: 6, A7: 7, A8: 8, A9: 9, A10: 10, C1: 3, C2: 6 },
    expected: 3,
    source: "pinned",
    note: "pinned: the same data and bins in a plain cell (no array range): the array result {3;3;4} shows its top-left element, 3 (SPEC-127 legacy array functions in a plain cell; SPEC-120 \"Implicit intersection\").",
  },
  {
    formula: "FREQUENCY(A1:A5)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: wrong argument count (1, FREQUENCY needs data_array, bins_array); SPEC-120 Errors: a wrong argument count is #VALUE!.",
  },

  // Pinned rows lie on the exact line y = 2x — (1,2), (2,4), (3,6), (4,8) —
  // where CORREL, PEARSON and RSQ are 1, SLOPE 2, INTERCEPT and STEYX 0 and
  // FORECAST at 4 is 8. CORREL's worked example is a screenshot, so its
  // Excel rows come from its Remarks. FORECAST and FORECAST.LINEAR share
  // one page, which states their syntax and usage are the same.
  // --- CORREL: array1, array2 --------------------------------------------
  {
    formula: "CORREL(A1:A3,B1:B3)",
    cells: { A1: 1, A2: 2, A3: 3, B1: 2, B2: 4, B3: 6 },
    expected: 1,
    source: "pinned",
    note: "pinned: array1 (A1:A3) and array2 (B1:B3) lie exactly on the line y=2x, so every deviation pairs as (dx,2dx); s12=2·s11 and s22=4·s11, giving r=s12/√(s11·s22)=1 exactly.",
  },
  {
    formula: "CORREL(A1,B1:B2)",
    cells: { A1: 1, B1: 2, B2: 3 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'CORREL Remarks, verbatim: "If array1 and array2 have a different number of data points, CORREL returns a #N/A error." (array1 has 1 point, array2 has 2).',
  },
  {
    formula: "CORREL(A1,B1)",
    cells: { A1: 5, B1: 1 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'CORREL Remarks, verbatim: "If either array1 or array2 is empty, or if s (the standard deviation) of their values equals zero, CORREL returns a #DIV/0! error." (a single point deviates by 0 from its own mean, so s=0 for both arrays).',
  },

  // --- PEARSON: array1, array2 --------------------------------------------
  {
    formula: "PEARSON(A3:A7,B3:B7)",
    cells: { A3: 9, A4: 7, A5: 5, A6: 3, A7: 1, B3: 10, B4: 6, B5: 1, B6: 5, B7: 3 },
    expected: 0.699379,
    digits: 6,
    source: "Excel",
    note: "PEARSON function, example: Pearson's r for the independent values in A3:A7 (below the sheet's \"Data\"/\"Independent values\" header rows 1-2) and dependent values in B3:B7; Microsoft prints 0.699379.",
  },
  {
    formula: "PEARSON(A1:A3,B1:B3)",
    cells: { A1: 1, A2: 2, A3: 3, B1: 2, B2: 4, B3: 6 },
    expected: 1,
    source: "pinned",
    note: "pinned: A1:A3 and B1:B3 lie exactly on y=2x (as CORREL above; Pearson's r and CORREL share one formula), so r=1 exactly.",
  },
  {
    formula: "PEARSON(A1,B1:B2)",
    cells: { A1: 1, B1: 2, B2: 3 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'PEARSON Remarks, verbatim: "If array1 and array2 are empty or have a different number of data points, PEARSON returns the #N/A error value." (array1 has 1 point, array2 has 2).',
  },

  // --- RSQ: known_y's, known_x's ------------------------------------------
  {
    formula: "RSQ(A3:A9,B3:B9)",
    cells: { A3: 2, A4: 3, A5: 9, A6: 1, A7: 8, A8: 7, A9: 5, B3: 6, B4: 5, B5: 11, B6: 7, B7: 5, B8: 4, B9: 4 },
    expected: 0.05795,
    digits: 4,
    source: "Excel",
    note: "RSQ function, example: the square of Pearson's r through known y's A3:A9 and known x's B3:B9 (below the sheet's \"Data\"/\"Known y\"/\"Known x\" header rows 1-2); Microsoft prints 0.05795.",
  },
  {
    formula: "RSQ(A1:A3,B1:B3)",
    cells: { A1: 2, A2: 4, A3: 6, B1: 1, B2: 2, B3: 3 },
    expected: 1,
    source: "pinned",
    note: "pinned: known y's A1:A3 and known x's B1:B3 lie exactly on y=2x, so r=1 and RSQ=r²=1 exactly.",
  },
  {
    formula: "RSQ(A1,B1)",
    cells: { A1: 2, B1: 5 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'RSQ Remarks, verbatim: "If known_y\'s and known_x\'s contain only 1 data point, RSQ returns the #DIV/0! error value."',
  },

  // --- COVAR (legacy COVARIANCE.P): array1, array2 ------------------------
  {
    formula: "COVAR(A2:A6,B2:B6)",
    cells: { A2: 3, A3: 2, A4: 4, A5: 5, A6: 6, B2: 9, B3: 7, B4: 12, B5: 15, B6: 17 },
    expected: 5.2,
    source: "Excel",
    note: "COVAR function, example: population covariance of Data1 in A2:A6 and Data2 in B2:B6; Microsoft prints 5.2 (exact: Σ(dx·dy)=26 over n=5 pairs, deviations from means 4 and 12).",
  },
  {
    formula: "COVAR(A1:A4,B1:B4)",
    cells: { A1: 1, A2: 2, A3: 3, A4: 4, B1: 2, B2: 4, B3: 6, B4: 8 },
    expected: 2.5,
    source: "pinned",
    note: "pinned: A1:A4 and B1:B4 lie on y=2x; deviations dx=-1.5,-0.5,0.5,1.5 and dy=2·dx, so Σ(dx·dy)=2·Σdx²=2·5=10, population covariance=10/4=2.5 exactly.",
  },
  {
    formula: "COVAR(A1,B1)",
    cells: { A1: 5 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'COVAR Remarks, verbatim: "If either array1 or array2 is empty, COVAR returns the #DIV/0! error value." (B1 left blank).',
  },

  // --- COVARIANCE.P: array1, array2 ---------------------------------------
  {
    formula: "COVARIANCE.P(A2:A6,B2:B6)",
    cells: { A2: 3, A3: 2, A4: 4, A5: 5, A6: 6, B2: 9, B3: 7, B4: 12, B5: 15, B6: 17 },
    expected: 5.2,
    source: "Excel",
    note: "COVARIANCE.P function, example: population covariance of Data1 in A2:A6 and Data2 in B2:B6 (the same pairs as COVAR's own page); Microsoft prints 5.2.",
  },
  {
    formula: "COVARIANCE.P(A1:A4,B1:B4)",
    cells: { A1: 1, A2: 2, A3: 3, A4: 4, B1: 2, B2: 4, B3: 6, B4: 8 },
    expected: 2.5,
    source: "pinned",
    note: "pinned: as COVAR above, A1:A4/B1:B4 lie on y=2x giving population covariance 10/4=2.5 exactly.",
  },
  {
    formula: "COVARIANCE.P(A1,B1)",
    cells: { A1: 5 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'COVARIANCE.P Remarks, verbatim: "If either array1 or array2 is empty, COVARIANCE.P returns the #DIV/0! error value." (B1 left blank).',
  },

  // --- COVARIANCE.S: array1, array2 ---------------------------------------
  {
    formula: "COVARIANCE.S({2,4,8},{5,11,12})",
    expected: 9.666666667,
    digits: 10,
    source: "Excel",
    note: "COVARIANCE.S function, example 1: sample covariance of the two arrays entered as constants, exactly as the page writes the formula; Microsoft prints 9.666666667.",
  },
  {
    formula: "COVARIANCE.S(A3:A5,B3:B5)",
    cells: { A3: 2, A4: 4, A5: 8, B3: 5, B4: 11, B5: 12 },
    expected: 9.666666667,
    digits: 10,
    source: "Excel",
    note: "COVARIANCE.S function, example 2: the identical data entered as cell ranges A3:A5 and B3:B5 instead of array constants; Microsoft prints the same 9.666666667.",
  },
  {
    formula: "COVARIANCE.S(A1:A3,B1:B3)",
    cells: { A1: 1, A2: 2, A3: 3, B1: 2, B2: 4, B3: 6 },
    expected: 2,
    source: "pinned",
    note: "pinned: A1:A3/B1:B3 lie on y=2x; Σ(dx·dy)=4 (dx=-1,0,1 and dy=2·dx), sample covariance=4/(3-1)=2 exactly.",
  },
  {
    formula: "COVARIANCE.S(A1,B1)",
    cells: { A1: 3, B1: 5 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'COVARIANCE.S Remarks, verbatim: "If either array1 or array2 is empty or contains only 1 data point each, COVARIANCE.S returns the #DIV/0! error value."',
  },

  // --- SLOPE: known_y's, known_x's ----------------------------------------
  {
    formula: "SLOPE(A3:A9,B3:B9)",
    cells: { A3: 2, A4: 3, A5: 9, A6: 1, A7: 8, A8: 7, A9: 5, B3: 6, B4: 5, B5: 11, B6: 7, B7: 5, B8: 4, B9: 4 },
    expected: 0.305556,
    digits: 6,
    source: "Excel",
    note: "SLOPE function, example: slope of the least-squares line through known y's A3:A9 and known x's B3:B9; Microsoft prints 0.305556.",
  },
  {
    formula: "SLOPE(A1:A3,B1:B3)",
    cells: { A1: 2, A2: 4, A3: 6, B1: 1, B2: 2, B3: 3 },
    expected: 2,
    source: "pinned",
    note: "pinned: known y's A1:A3 and known x's B1:B3 lie exactly on y=2x, so the regression slope is exactly 2.",
  },
  {
    formula: "SLOPE(A1,B1:B2)",
    cells: { A1: 2, B1: 1, B2: 2 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'SLOPE Remarks, verbatim: "If known_y\'s and known_x\'s are empty or have a different number of data points, SLOPE returns the #N/A error value." (1 known y against 2 known x\'s).',
  },

  // --- INTERCEPT: known_y's, known_x's ------------------------------------
  {
    formula: "INTERCEPT(A2:A6,B2:B6)",
    cells: { A2: 2, A3: 3, A4: 9, A5: 1, A6: 8, B2: 6, B3: 5, B4: 11, B5: 7, B6: 5 },
    expected: 0.0483871,
    digits: 6,
    source: "Excel",
    note: "INTERCEPT function, example: where the least-squares line through known y's A2:A6 and known x's B2:B6 crosses the y-axis; Microsoft prints 0.0483871.",
  },
  {
    formula: "INTERCEPT(A1:A3,B1:B3)",
    cells: { A1: 2, A2: 4, A3: 6, B1: 1, B2: 2, B3: 3 },
    expected: 0,
    source: "pinned",
    note: "pinned: known y's A1:A3 and known x's B1:B3 lie on y=2x, a line through the origin, so the intercept is exactly 0.",
  },
  {
    formula: "INTERCEPT(A1,B1:B2)",
    cells: { A1: 2, B1: 1, B2: 2 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'INTERCEPT Remarks, verbatim: "If known_y\'s and known_x\'s contain a different number of data points or contain no data points, INTERCEPT returns the #N/A error value." (1 known y against 2 known x\'s).',
  },

  // --- STEYX: known_y's, known_x's ----------------------------------------
  {
    formula: "STEYX(A3:A9,B3:B9)",
    cells: { A3: 2, A4: 3, A5: 9, A6: 1, A7: 8, A8: 7, A9: 5, B3: 6, B4: 5, B5: 11, B6: 7, B7: 5, B8: 4, B9: 4 },
    expected: 3.305719,
    digits: 7,
    source: "Excel",
    note: "STEYX function, example: standard error of the y's predicted from known y's A3:A9 and known x's B3:B9; Microsoft prints 3.305719.",
  },
  {
    formula: "STEYX(A1:A3,B1:B3)",
    cells: { A1: 2, A2: 4, A3: 6, B1: 1, B2: 2, B3: 3 },
    expected: 0,
    source: "pinned",
    note: "pinned: known y's A1:A3 and known x's B1:B3 lie exactly on y=2x, so every predicted y matches the actual y and the standard error is exactly 0.",
  },
  {
    formula: "STEYX(A1:A2,B1:B2)",
    cells: { A1: 1, A2: 2, B1: 3, B2: 4 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'STEYX Remarks, verbatim: "If known_y\'s and known_x\'s are empty or have less than three data points, STEYX returns the #DIV/0! error value." (2 points).',
  },

  // --- FORECAST (legacy FORECAST.LINEAR): x, known_y's, known_x's --------
  {
    formula: "FORECAST(30,A2:A6,B2:B6)",
    cells: { A2: 6, A3: 7, A4: 9, A5: 15, A6: 21, B2: 20, B3: 28, B4: 31, B5: 38, B6: 40 },
    expected: 10.607253,
    digits: 8,
    source: "Excel",
    note: 'FORECAST function, example: the "FORECAST and FORECAST.LINEAR functions" page (both share it) states "The syntax and usage of the two functions are the same" and gives this worked example under the FORECAST.LINEAR name — predicted y at x=30 from known y\'s A2:A6 and known x\'s B2:B6; Microsoft prints 10.607253, equally valid for legacy FORECAST.',
  },
  {
    formula: "FORECAST(4,A1:A3,B1:B3)",
    cells: { A1: 2, A2: 4, A3: 6, B1: 1, B2: 2, B3: 3 },
    expected: 8,
    source: "pinned",
    note: "pinned: known y's A1:A3 and known x's B1:B3 lie exactly on y=2x, so the forecast at x=4 is exactly 8.",
  },
  {
    formula: "FORECAST(5,A1:A3,B1:B3)",
    cells: { A1: 1, A2: 2, A3: 3, B1: 4, B2: 4, B3: 4 },
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'FORECAST and FORECAST.LINEAR Remarks, verbatim condition: "If the variance of known_x\'s equals zero" (listed against the #DIV/0! error) — known x\'s B1:B3 are all 4.',
  },

  // --- FORECAST.LINEAR: x, known_y's, known_x's ---------------------------
  {
    formula: "FORECAST.LINEAR(30,A2:A6,B2:B6)",
    cells: { A2: 6, A3: 7, A4: 9, A5: 15, A6: 21, B2: 20, B3: 28, B4: 31, B5: 38, B6: 40 },
    expected: 10.607253,
    digits: 8,
    source: "Excel",
    note: "FORECAST.LINEAR function, example: predicted y at x=30 from known y's A2:A6 and known x's B2:B6; Microsoft prints 10.607253.",
  },
  {
    formula: "FORECAST.LINEAR(4,A1:A3,B1:B3)",
    cells: { A1: 2, A2: 4, A3: 6, B1: 1, B2: 2, B3: 3 },
    expected: 8,
    source: "pinned",
    note: "pinned: as FORECAST above, known y's/x's on y=2x give a forecast of exactly 8 at x=4.",
  },
  {
    formula: "FORECAST.LINEAR(5,A1:A3,B1:B2)",
    cells: { A1: 1, A2: 2, A3: 3, B1: 4, B2: 5 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'FORECAST and FORECAST.LINEAR Remarks, verbatim condition: "If known_y\'s or known_x\'s is empty or one has more data points than the other" (listed against the #N/A error) — 3 known y\'s against 2 known x\'s.',
  },

  // --- STANDARDIZE: x, mean, standard_dev ---------------------------------
  {
    formula: "STANDARDIZE(A2,A3,A4)",
    cells: { A2: 42, A3: 40, A4: 1.5 },
    expected: 1.33333333,
    digits: 9,
    source: "Excel",
    note: "STANDARDIZE function, example: how many standard deviations 42 (A2) lies from a mean of 40 (A3) with standard deviation 1.5 (A4); Microsoft prints 1.33333333.",
  },
  {
    formula: "STANDARDIZE(10,0,5)",
    expected: 2,
    source: "pinned",
    note: "pinned: (10-0)/5 = 2 exactly.",
  },
  {
    formula: "STANDARDIZE(42,40,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'STANDARDIZE Remarks, verbatim: "If standard_dev ≤ 0, STANDARDIZE returns the #NUM! error value." (standard_dev 0).',
  },

  // --- KURT: number1, [number2], … ----------------------------------------
  {
    formula: "KURT(A2:A11)",
    cells: { A2: 3, A3: 4, A4: 5, A5: 2, A6: 3, A7: 4, A8: 5, A9: 6, A10: 4, A11: 7 },
    expected: -0.151799637,
    digits: 9,
    source: "Excel",
    note: "KURT function, example: kurtosis of the data set in A2:A11; Microsoft prints -0.151799637.",
  },
  {
    formula: "KURT(1,-1,1,-1)",
    expected: -6,
    source: "pinned",
    note: "pinned: 4 values ±1 about mean 0; sample variance=4/3 so s⁴=16/9, Σ((x-mean)/s)⁴=4÷(16/9)=9/4; KURT=(n(n+1)/((n-1)(n-2)(n-3)))·(9/4)-3(n-1)²/((n-2)(n-3))=(20/6)(9/4)-27/2=7.5-13.5=-6 exactly.",
  },
  {
    formula: "KURT(1,2,3)",
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'KURT Remarks, verbatim: "If there are fewer than four data points, or if the standard deviation of the sample equals zero, KURT returns the #DIV/0! error value." (3 points).',
  },

  // --- SKEW: number1, [number2], … ----------------------------------------
  {
    formula: "SKEW(A2:A11)",
    cells: { A2: 3, A3: 4, A4: 5, A5: 2, A6: 3, A7: 4, A8: 5, A9: 6, A10: 4, A11: 7 },
    expected: 0.359543,
    digits: 6,
    source: "Excel",
    note: "SKEW function, example: skewness of the data set in A2:A11; Microsoft prints 0.359543.",
  },
  {
    formula: "SKEW(1,2,3)",
    expected: 0,
    source: "pinned",
    note: "pinned: 1,2,3 is symmetric about its mean 2 (deviations -1,0,1), so the sum of cubed deviations is 0 and SKEW is exactly 0.",
  },
  {
    formula: "SKEW(1,2)",
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'SKEW Remarks, verbatim: "If there are fewer than three data points, or the sample standard deviation is zero, SKEW returns the #DIV/0! error value." (2 points).',
  },

  // --- SKEW.P: number1, [number2], … --------------------------------------
  {
    formula: "SKEW.P(A2:A11)",
    cells: { A2: 3, A3: 4, A4: 5, A5: 2, A6: 3, A7: 4, A8: 5, A9: 6, A10: 4, A11: 7 },
    expected: 0.303193,
    digits: 6,
    source: "Excel",
    note: "SKEW.P function, example: population skewness of the data set in A2:A11; Microsoft prints 0.303193.",
  },
  {
    formula: "SKEW.P(1,2,4,5)",
    expected: 0,
    source: "pinned",
    note: "pinned: 1,2,4,5 is symmetric about its mean 3 (deviations -2,-1,1,2), so the sum of cubed deviations is 0 and SKEW.P is exactly 0.",
  },
  {
    formula: "SKEW.P(5)",
    expected: { error: "#DIV/0!" },
    source: "Excel",
    note: 'SKEW.P Remarks, verbatim: "If there are fewer than three data points, or the sample standard deviation is zero, SKEW.P returns the #DIV/0! error value." (1 point, trivially zero deviation from its own mean).',
  },

  // Pinned rows rest on identities: a normal CDF at its own mean is 0.5, so
  // the inverse of 0.5 is the mean; GAMMA(n) = (n−1)! for a positive
  // integer; GAMMALN(1) = GAMMALN(2) = ln 1 = 0.
  // --- NORM.DIST: x, mean, standard_dev, cumulative -------------------------
  {
    formula: "NORM.DIST(A2,A3,A4,TRUE)",
    cells: { A2: 42, A3: 40, A4: 1.5 },
    expected: 0.9087888,
    digits: 7,
    source: "Excel",
    note: "NORM.DIST function, example: cumulative distribution at x=42, mean 40, sd 1.5; Microsoft prints 0.9087888.",
  },
  {
    formula: "NORM.DIST(A2,A3,A4,FALSE)",
    cells: { A2: 42, A3: 40, A4: 1.5 },
    expected: 0.10934,
    digits: 5,
    source: "Excel",
    note: "NORM.DIST function, example: probability density at the same x=42, mean 40, sd 1.5; Microsoft prints 0.10934.",
  },
  {
    formula: "NORM.DIST(42,40,0,TRUE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'NORM.DIST function Remarks: "If standard_dev ≤ 0, NORM.DIST returns the #NUM! error value."',
  },
  {
    formula: "NORM.DIST(40,40,1.5,TRUE)",
    expected: 0.5,
    source: "pinned",
    note: "pinned: at x = mean, z = (x−mean)/sd = 0, and the standard normal CDF at 0 is exactly 0.5 by symmetry whatever sd is; NORM.DIST(mean,mean,sd,TRUE) = 0.5 exactly.",
  },

  // --- NORM.INV: probability, mean, standard_dev ----------------------------
  {
    formula: "NORM.INV(A2,A3,A4)",
    cells: { A2: 0.908789, A3: 40, A4: 1.5 },
    expected: 42.000002,
    digits: 8,
    source: "Excel",
    note: "NORM.INV function, example: the x at which the cumulative distribution reaches 0.908789, mean 40, sd 1.5; Microsoft prints 42.000002.",
  },
  {
    formula: "NORM.INV(0,40,1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'NORM.INV function Remarks: "If probability <= 0 or if probability >= 1, NORM.INV returns the #NUM! error value."',
  },
  {
    formula: 'NORM.INV("x",40,1.5)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'NORM.INV function Remarks: "If any argument is nonnumeric, NORM.INV returns the #VALUE! error value."',
  },
  {
    formula: "NORM.INV(0.5,40,1.5)",
    expected: 40,
    source: "pinned",
    note: "pinned: NORM.INV is NORM.DIST's inverse; a normal CDF at its own mean is exactly 0.5 (symmetry), so NORM.INV(0.5,mean,sd) = mean exactly.",
  },

  // --- NORM.S.DIST: z, cumulative --------------------------------------------
  {
    formula: "NORM.S.DIST(1.333333,TRUE)",
    expected: 0.908788726,
    digits: 9,
    source: "Excel",
    note: "NORM.S.DIST function, example: standard normal cumulative distribution at z=1.333333; Microsoft prints 0.908788726.",
  },
  {
    formula: "NORM.S.DIST(1.333333,FALSE)",
    expected: 0.164010148,
    digits: 9,
    source: "Excel",
    note: "NORM.S.DIST function, example: standard normal density at the same z=1.333333; Microsoft prints 0.164010148.",
  },
  {
    formula: 'NORM.S.DIST("x",TRUE)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'NORM.S.DIST function Remarks: "If z is nonnumeric, NORM.S.DIST returns the #VALUE! error."',
  },
  {
    formula: "NORM.S.DIST(0,TRUE)",
    expected: 0.5,
    source: "pinned",
    note: "pinned: the standard normal density is symmetric about 0, so its cumulative distribution at 0 is exactly 0.5.",
  },

  // --- NORM.S.INV: probability ------------------------------------------------
  {
    formula: "NORM.S.INV(0.908789)",
    expected: 1.3333347,
    digits: 8,
    source: "Excel",
    note: "NORM.S.INV function, example: the z at which the standard normal cumulative distribution reaches 0.908789; Microsoft prints 1.3333347.",
  },
  {
    formula: "NORM.S.INV(0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'NORM.S.INV function Remarks: "If probability <= 0 or if probability >= 1, NORMS.INV returns the #NUM! error value."',
  },
  {
    formula: 'NORM.S.INV("x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'NORM.S.INV function Remarks: "If probability is nonnumeric, NORMS.INV returns the #VALUE! error value."',
  },
  {
    formula: "NORM.S.INV(0.5)",
    expected: 0,
    source: "pinned",
    note: "pinned: NORM.S.DIST(0,TRUE) = 0.5 exactly (0 is the standard normal's median), so its inverse NORM.S.INV(0.5) = 0 exactly.",
  },

  // --- NORMDIST: x, mean, standard_dev, cumulative (own page, legacy NORM.DIST) --
  {
    formula: "NORMDIST(A2,A3,A4,TRUE)",
    cells: { A2: 42, A3: 40, A4: 1.5 },
    expected: 0.9087888,
    digits: 7,
    source: "Excel",
    note: "NORMDIST function, example: cumulative distribution at x=42, mean 40, sd 1.5; Microsoft prints 0.9087888.",
  },
  {
    formula: "NORMDIST(A2,A3,A4,FALSE)",
    cells: { A2: 42, A3: 40, A4: 1.5 },
    expected: 0.10934,
    digits: 5,
    source: "Excel",
    note: "NORMDIST function, example: probability density at the same x=42, mean 40, sd 1.5; Microsoft prints 0.10934.",
  },
  {
    formula: "NORMDIST(42,40,0,TRUE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'NORMDIST function Remarks: "If standard_dev ≤ 0, NORMDIST returns the #NUM! error value."',
  },
  {
    formula: "NORMDIST(40,40,1.5,TRUE)",
    expected: 0.5,
    source: "pinned",
    note: "pinned: a normal CDF at its own mean is exactly 0.5 by symmetry; NORMDIST(mean,mean,sd,TRUE) = 0.5 exactly.",
  },

  // --- NORMINV: probability, mean, standard_dev (own page, legacy NORM.INV) -----
  {
    formula: "NORMINV(A2,A3,A4)",
    cells: { A2: 0.908789, A3: 40, A4: 1.5 },
    expected: 42.000002,
    digits: 8,
    source: "Excel",
    note: "NORMINV function, example: the x at which the cumulative distribution reaches 0.908789, mean 40, sd 1.5; Microsoft prints 42.000002.",
  },
  {
    formula: "NORMINV(0,40,1.5)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'NORMINV function Remarks: "If probability <= 0 or if probability >= 1, NORMINV returns the #NUM! error value."',
  },
  {
    formula: "NORMINV(0.5,40,1.5)",
    expected: 40,
    source: "pinned",
    note: "pinned: NORMINV is NORMDIST's inverse; a normal CDF at its own mean is exactly 0.5 (symmetry), so NORMINV(0.5,mean,sd) = mean exactly.",
  },

  // --- NORMSDIST: z (own page, legacy NORM.S.DIST, always cumulative) -----------
  {
    formula: "NORMSDIST(1.333333)",
    expected: 0.908788726,
    digits: 9,
    source: "Excel",
    note: "NORMSDIST function, example: standard normal cumulative distribution at z=1.333333; Microsoft prints 0.908788726.",
  },
  {
    formula: 'NORMSDIST("x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'NORMSDIST function Remarks: "If z is nonnumeric, NORMSDIST returns the #VALUE! error value."',
  },
  {
    formula: "NORMSDIST(0)",
    expected: 0.5,
    source: "pinned",
    note: "pinned: NORMSDIST is the standard normal CDF; symmetric about 0, so NORMSDIST(0) = 0.5 exactly.",
  },

  // --- NORMSINV: probability (own page, legacy NORM.S.INV) -----------------------
  {
    formula: "NORMSINV(0.9088)",
    expected: 1.3334,
    digits: 5,
    source: "Excel",
    note: "NORMSINV function, example: the z at which NORMSDIST reaches 0.9088; Microsoft prints 1.3334.",
  },
  {
    formula: "NORMSINV(0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'NORMSINV function Remarks: "If Probability <= 0 or if Probability >= 1, NORMSINV returns the #NUM! error value."',
  },
  {
    formula: "NORMSINV(0.5)",
    expected: 0,
    source: "pinned",
    note: "pinned: NORMSDIST(0) = 0.5 exactly, so its inverse NORMSINV(0.5) = 0 exactly.",
  },

  // --- CONFIDENCE: alpha, standard_dev, size (own page, legacy CONFIDENCE.NORM) --
  {
    formula: "CONFIDENCE(A2,A3,A4)",
    cells: { A2: 0.05, A3: 2.5, A4: 50 },
    expected: 0.692951912,
    digits: 9,
    source: "Excel",
    note: "CONFIDENCE function, example: half-width of a confidence interval for a mean, significance 0.05, sd 2.5, n=50; Microsoft prints 0.692951912.",
  },
  {
    formula: "CONFIDENCE(0,2.5,50)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'CONFIDENCE function Remarks: "If Alpha is ≤ 0 or ≥ 1, CONFIDENCE returns the #NUM! error value."',
  },
  {
    formula: "CONFIDENCE(0.05,2.5,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'CONFIDENCE function Remarks: "If Size < 1, CONFIDENCE returns the #NUM! error value."',
  },
  {
    formula: 'CONFIDENCE("x",2.5,50)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'CONFIDENCE function Remarks: "If any argument is non-numeric, CONFIDENCE returns the #VALUE! error value."',
  },

  // --- CONFIDENCE.NORM: alpha, standard_dev, size ----------------------------
  {
    formula: "CONFIDENCE.NORM(A2,A3,A4)",
    cells: { A2: 0.05, A3: 2.5, A4: 50 },
    expected: 0.692952,
    digits: 6,
    source: "Excel",
    note: "CONFIDENCE.NORM function, example: half-width of a confidence interval for a mean, significance 0.05, sd 2.5, n=50; Microsoft prints 0.692952.",
  },
  {
    formula: "CONFIDENCE.NORM(1,2.5,50)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'CONFIDENCE.NORM function Remarks: "If alpha ≤ 0 or alpha ≥ 1, CONFIDENCE.NORM returns the #NUM! error value."',
  },
  {
    formula: "CONFIDENCE.NORM(0.05,0,50)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'CONFIDENCE.NORM function Remarks: "If standard_dev ≤ 0, CONFIDENCE.NORM returns the #NUM! error value."',
  },
  {
    formula: 'CONFIDENCE.NORM(0.05,2.5,"x")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'CONFIDENCE.NORM function Remarks: "If any argument is nonnumeric, CONFIDENCE.NORM returns the #VALUE! error value."',
  },

  // --- GAMMA: number -----------------------------------------------------------
  {
    formula: "GAMMA(2.5)",
    expected: 1.329,
    digits: 4,
    source: "Excel",
    note: "GAMMA function, example, row 1: the gamma function at 2.5; Microsoft prints 1.329.",
  },
  {
    formula: "GAMMA(-3.75)",
    expected: 0.268,
    digits: 3,
    source: "Excel",
    note: "GAMMA function, example, row 2: the gamma function at -3.75; Microsoft prints 0.268.",
  },
  {
    formula: "GAMMA(0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'GAMMA function, example, row 3, and Remarks: "If Number is a negative integer or 0, GAMMA returns the #NUM! error value."',
  },
  {
    formula: "GAMMA(-2)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'GAMMA function, example, row 4: a negative integer is not a valid argument either, per the same Remarks sentence as GAMMA(0).',
  },
  {
    formula: "GAMMA(5)",
    expected: 24,
    source: "pinned",
    note: "pinned: GAMMA(n) = (n−1)! for a positive integer n (Remarks' recurrence Γ(N+1) = N·Γ(N), unwound from Γ(1) = 1); GAMMA(5) = 4! = 24 exactly.",
  },

  // --- GAMMALN: x ----------------------------------------------------------------
  {
    formula: "GAMMALN(4)",
    expected: 1.7917595,
    digits: 8,
    source: "Excel",
    note: "GAMMALN function, example: the natural log of the gamma function at 4; Microsoft prints 1.7917595.",
  },
  {
    formula: "GAMMALN(0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'GAMMALN function Remarks: "If x ≤ 0, GAMMALN returns the #NUM! error value."',
  },
  {
    formula: "GAMMALN(1)",
    expected: 0,
    source: "pinned",
    note: "pinned: Γ(1) = 0! = 1 and ln(1) = 0; GAMMALN(1) = 0 exactly.",
  },
  {
    formula: "GAMMALN(2)",
    expected: 0,
    source: "pinned",
    note: "pinned: Γ(2) = 1! = 1 and ln(1) = 0; GAMMALN(2) = 0 exactly.",
  },

  // --- GAMMALN.PRECISE: x ---------------------------------------------------------
  {
    formula: "GAMMALN.PRECISE(4)",
    expected: 1.7917595,
    digits: 8,
    source: "Excel",
    note: "GAMMALN.PRECISE function, example: the natural log of the gamma function at 4; Microsoft prints 1.7917595.",
  },
  {
    formula: "GAMMALN.PRECISE(0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'GAMMALN.PRECISE function Remarks: "If x ≤ 0, GAMMALN.PRECISE returns the #NUM! error value."',
  },
  {
    formula: "GAMMALN.PRECISE(1)",
    expected: 0,
    source: "pinned",
    note: "pinned: Γ(1) = 0! = 1 and ln(1) = 0; GAMMALN.PRECISE(1) = 0 exactly.",
  },
  {
    formula: "GAMMALN.PRECISE(2)",
    expected: 0,
    source: "pinned",
    note: "pinned: Γ(2) = 1! = 1 and ln(1) = 0; GAMMALN.PRECISE(2) = 0 exactly.",
  },

  // Pinned rows rest on identities: the t distribution is symmetric about
  // 0 (CDF 0.5 there, both tails 1, the left inverse of 0.5 is 0); a Poisson
  // of mean 0 is all at 0; the exponential CDF at 0 is 1 − e⁰ = 0.
  // --- T.DIST: x, deg_freedom, cumulative ---------------------------------
  {
    formula: "T.DIST(60,1,TRUE)",
    expected: 0.99469533,
    digits: 8,
    source: "Excel",
    note: "T.DIST function, example row 1: left-tailed cumulative distribution at x=60 with 1 degree of freedom; Microsoft prints 0.99469533.",
  },
  {
    formula: "T.DIST(8,3,FALSE)",
    expected: 0.00073691,
    digits: 5,
    source: "Excel",
    note: "T.DIST function, example row 2: left-tailed probability density at x=8 with 3 degrees of freedom; Microsoft prints 0.00073691.",
  },
  {
    formula: "T.DIST(0,10,TRUE)",
    expected: 0.5,
    digits: 15,
    source: "pinned",
    note: "pinned: Student's t-distribution is symmetric about 0, so its left-tailed CDF at x=0 is exactly 0.5 for any degrees of freedom (SPEC-127 pin: T.DIST(0,n,TRUE)=0.5).",
  },
  {
    formula: 'T.DIST("a",1,TRUE)',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'T.DIST function, Remarks: "If any argument is nonnumeric, T.DIST returns the #VALUE! error value." — x is text.',
  },

  // --- T.DIST.2T: x, deg_freedom -------------------------------------------
  {
    formula: "T.DIST.2T(A2,A3)",
    cells: { A2: 1.959999998, A3: 60 },
    expected: 0.054645,
    digits: 5,
    source: "Excel",
    note: "T.DIST.2T function, example: two-tailed distribution for the value and degrees of freedom in A2:A3; Microsoft's Result column shows 5.46% but the Description states the more precise 0.054645, used here.",
  },
  {
    formula: "T.DIST.2T(0,10)",
    expected: 1,
    digits: 15,
    source: "pinned",
    note: "pinned: P(|T|>0) excludes only the single point T=0, which has probability 0 in a continuous distribution, so T.DIST.2T(0,n) is exactly 1 for any degrees of freedom (SPEC-127 pin).",
  },
  {
    formula: "T.DIST.2T(-1,10)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'T.DIST.2T function, Remarks: "If x < 0, then T.DIST.2T returns the #NUM! error value."',
  },

  // --- T.DIST.RT: x, deg_freedom --------------------------------------------
  {
    formula: "T.DIST.RT(A2,A3)",
    cells: { A2: 1.959999998, A3: 60 },
    expected: 0.027322,
    digits: 5,
    source: "Excel",
    note: "T.DIST.RT function, example: right-tailed distribution for the value and degrees of freedom in A2:A3; Microsoft prints 0.027322.",
  },
  {
    formula: "T.DIST.RT(0,10)",
    expected: 0.5,
    digits: 15,
    source: "pinned",
    note: "pinned: by the same symmetry as T.DIST, P(T>0) is exactly 0.5 for any degrees of freedom (SPEC-127 pin: T.DIST.RT(0,n)=0.5).",
  },
  {
    formula: "T.DIST.RT(1,0)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'T.DIST.RT function, Remarks: "If deg_freedom < 1, T.DIST.RT returns the #NUM! error value." — deg_freedom=0.',
  },

  // --- T.INV: probability, deg_freedom --------------------------------------
  {
    formula: "T.INV(0.75,2)",
    expected: 0.8164966,
    digits: 7,
    source: "Excel",
    note: "T.INV function, example: left-tailed inverse at probability 75% with 2 degrees of freedom; Microsoft prints 0.8164966.",
  },
  {
    formula: "T.INV(0.5,10)",
    expected: 0,
    digits: 15,
    source: "pinned",
    note: "pinned: the left-tailed inverse at probability 0.5 is the median of a distribution symmetric about 0, exactly 0 for any degrees of freedom (SPEC-127 pin: T.INV(0.5,n)=0).",
  },
  {
    formula: "T.INV(0,10)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'T.INV function, Remarks: "If probability <= 0 or if probability > 1, T.INV returns the #NUM! error value."',
  },

  // --- T.INV.2T: probability, deg_freedom -----------------------------------
  {
    formula: "T.INV.2T(A2,A3)",
    cells: { A2: 0.546449, A3: 60 },
    expected: 0.606533,
    digits: 6,
    source: "Excel",
    note: "T.INV.2T function, example: two-tailed inverse for the probability and degrees of freedom in A2:A3; Microsoft prints 0.606533.",
  },
  {
    formula: "T.INV.2T(1,10)",
    expected: 0,
    digits: 15,
    source: "pinned",
    note: "pinned: T.INV.2T solves P(|T|>t)=probability; at probability=1 that holds at t=0 (only the single point T=0, probability 0, is excluded) — the dual of T.DIST.2T(0,n)=1.",
  },
  {
    formula: "T.INV.2T(0,10)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'T.INV.2T function, Remarks: "if probability <= 0 or if probability > 1, T.INV.2T returns the #NUM! error value."',
  },

  // --- TDIST (legacy): x, deg_freedom, tails --------------------------------
  {
    formula: "TDIST(A2,A3,2)",
    cells: { A2: 1.959999998, A3: 60 },
    expected: 0.05464493,
    digits: 8,
    source: "Excel",
    note: "TDIST function, example row 1: two-tailed distribution for the value and degrees of freedom in A2:A3; Microsoft's Result column shows 5.46% but the Description states the more precise 0.054644930, used here.",
  },
  {
    formula: "TDIST(A2,A3,1)",
    cells: { A2: 1.959999998, A3: 60 },
    expected: 0.027322465,
    digits: 8,
    source: "Excel",
    note: "TDIST function, example row 2: one-tailed distribution, same A2:A3; Microsoft's Result column shows 2.73% but the Description states the more precise 0.027322465, used here.",
  },
  {
    formula: "TDIST(0,10,1)",
    expected: 0.5,
    digits: 15,
    source: "pinned",
    note: "pinned: tails=1 is TDIST's one-tailed (right-tail) distribution, Microsoft's own description; by symmetry P(T>0) is exactly 0.5 for any degrees of freedom.",
  },
  {
    formula: "TDIST(-1,60,2)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'TDIST function, Remarks: "If x < 0, then TDIST returns the #NUM! error value."',
  },

  // --- TINV (legacy): probability, deg_freedom ------------------------------
  {
    formula: "TINV(A2,A3)",
    cells: { A2: 0.05464, A3: 60 },
    expected: 1.96,
    digits: 3,
    source: "Excel",
    note: "TINV function, example: the t-value for the two-tailed probability and degrees of freedom in A2:A3; Microsoft prints 1.96.",
  },
  {
    formula: "TINV(1,10)",
    expected: 0,
    digits: 15,
    source: "pinned",
    note: "pinned: TINV is the two-tailed inverse (same rule as T.INV.2T); at probability=1 the solution is t=0, since only the single point T=0 (probability 0) is excluded from '|T|>0'.",
  },
  {
    formula: "TINV(0,60)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'TINV function, Remarks: "If probability <= 0 or if probability > 1, TINV returns the #NUM! error value."',
  },

  // --- BINOM.DIST: number_s, trials, probability_s, cumulative -------------
  {
    formula: "BINOM.DIST(A2,A3,A4,FALSE)",
    cells: { A2: 6, A3: 10, A4: 0.5 },
    expected: 0.2050781,
    digits: 7,
    source: "Excel",
    note: "BINOM.DIST function, example: probability of exactly 6 successes in 10 trials at probability_s=0.5; Microsoft prints 0.2050781.",
  },
  {
    formula: "BINOM.DIST(0,2,0.5,FALSE)",
    expected: 0.25,
    digits: 15,
    source: "pinned",
    note: "pinned: P(X=0) of Binomial(2,0.5) is C(2,0)*0.5^0*0.5^2 = 1*1*0.25 = 0.25 exactly (SPEC-127 pin).",
  },
  {
    formula: "BINOM.DIST(1,2,0.5,TRUE)",
    expected: 0.75,
    digits: 15,
    source: "pinned",
    note: "pinned: P(X<=1) of Binomial(2,0.5) is P(X=0)+P(X=1) = 0.25 + C(2,1)*0.5*0.5 = 0.25+0.5 = 0.75 exactly (SPEC-127 pin).",
  },
  {
    formula: "BINOM.DIST(3,2,0.5,FALSE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'BINOM.DIST function, Remarks: "If number_s < 0 or number_s > trials, BINOM.DIST returns the #NUM! error value." — number_s=3 > trials=2.',
  },

  // --- BINOMDIST (legacy): number_s, trials, probability_s, cumulative -----
  {
    formula: "BINOMDIST(A2,A3,A4,FALSE)",
    cells: { A2: 6, A3: 10, A4: 0.5 },
    expected: 0.2050781,
    digits: 7,
    source: "Excel",
    note: "BINOMDIST function, example: probability of exactly 6 successes in 10 trials at probability_s=0.5; Microsoft prints 0.2050781.",
  },
  {
    formula: "BINOMDIST(0,2,0.5,FALSE)",
    expected: 0.25,
    digits: 15,
    source: "pinned",
    note: "pinned: P(X=0) of Binomial(2,0.5) is C(2,0)*0.5^0*0.5^2 = 1*1*0.25 = 0.25 exactly, same rule as BINOM.DIST (SPEC-127 pin).",
  },
  {
    formula: "BINOMDIST(1,2,0.5,TRUE)",
    expected: 0.75,
    digits: 15,
    source: "pinned",
    note: "pinned: P(X<=1) of Binomial(2,0.5) is 0.25 + C(2,1)*0.5*0.5 = 0.25+0.5 = 0.75 exactly, same rule as BINOM.DIST (SPEC-127 pin).",
  },
  {
    formula: "BINOMDIST(3,2,0.5,FALSE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'BINOMDIST function, Remarks: "If number_s < 0 or number_s > trials, BINOMDIST returns the #NUM! error value." — number_s=3 > trials=2.',
  },

  // --- POISSON (legacy): x, mean, cumulative --------------------------------
  {
    formula: "POISSON(A2,A3,TRUE)",
    cells: { A2: 2, A3: 5 },
    expected: 0.124652,
    digits: 6,
    source: "Excel",
    note: "POISSON function, example row 1: cumulative probability of at most 2 events with mean 5; Microsoft prints 0.124652.",
  },
  {
    formula: "POISSON(A2,A3,FALSE)",
    cells: { A2: 2, A3: 5 },
    expected: 0.084224,
    digits: 5,
    source: "Excel",
    note: "POISSON function, example row 2: probability mass at exactly 2 events with mean 5; Microsoft prints 0.084224.",
  },
  {
    formula: "POISSON(0,0,TRUE)",
    expected: 1,
    digits: 15,
    source: "pinned",
    note: "pinned: with mean=0 the pmf's exponent on mean is 0 at x=0, and anything^0 is 1 by the formula's own convention, so e^-0 * 0^0/0! = 1*1/1 = 1 exactly; both TRUE and FALSE give 1 since all the mass sits at 0.",
  },
  {
    formula: "POISSON(2,-1,TRUE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'POISSON function, Remarks: "If mean < 0, POISSON returns the #NUM! error value."',
  },

  // --- POISSON.DIST: x, mean, cumulative ------------------------------------
  {
    formula: "POISSON.DIST(A2,A3,TRUE)",
    cells: { A2: 2, A3: 5 },
    expected: 0.124652,
    digits: 6,
    source: "Excel",
    note: "POISSON.DIST function, example row 1: cumulative probability of at most 2 events with mean 5; Microsoft prints 0.124652.",
  },
  {
    formula: "POISSON.DIST(A2,A3,FALSE)",
    cells: { A2: 2, A3: 5 },
    expected: 0.084224,
    digits: 5,
    source: "Excel",
    note: "POISSON.DIST function, example row 2: probability mass at exactly 2 events with mean 5; Microsoft prints 0.084224.",
  },
  {
    formula: "POISSON.DIST(0,0,TRUE)",
    expected: 1,
    digits: 15,
    source: "pinned",
    note: "pinned: with mean=0 the pmf's exponent on mean is 0 at x=0, and anything^0 is 1 by the formula's own convention, so e^-0 * 0^0/0! = 1*1/1 = 1 exactly (same rule as POISSON; SPEC-127 pin: POISSON.DIST(0,0,…)=1).",
  },
  {
    formula: "POISSON.DIST(-1,5,TRUE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'POISSON.DIST function, Remarks: "If x < 0, POISSON.DIST returns the #NUM! error value."',
  },

  // --- EXPON.DIST: x, lambda, cumulative ------------------------------------
  {
    formula: "EXPON.DIST(A2,A3,TRUE)",
    cells: { A2: 0.2, A3: 10 },
    expected: 0.86466472,
    digits: 8,
    source: "Excel",
    note: "EXPON.DIST function, example row 1: cumulative distribution at x=0.2 with lambda=10; Microsoft prints 0.86466472.",
  },
  {
    formula: "EXPON.DIST(0.2,10,FALSE)",
    expected: 1.35335283,
    digits: 9,
    source: "Excel",
    note: "EXPON.DIST function, example row 2: probability density at x=0.2 with lambda=10 (Microsoft's own formula uses the literal values, not A2:A3); Microsoft prints 1.35335283.",
  },
  {
    formula: "EXPON.DIST(0,1,TRUE)",
    expected: 0,
    digits: 15,
    source: "pinned",
    note: "pinned: the exponential CDF is 1-e^(-lambda*x); at x=0 that is 1-e^0 = 1-1 = 0 exactly, for any lambda (SPEC-127 pin: EXPON.DIST(0,1,TRUE)=0).",
  },
  {
    formula: "EXPON.DIST(1,0,TRUE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'EXPON.DIST function, Remarks: "If lambda <= 0, EXPON.DIST returns the #NUM! error value." — lambda=0.',
  },

  // --- EXPONDIST (legacy): x, lambda, cumulative ----------------------------
  {
    formula: "EXPONDIST(A2,A3,TRUE)",
    cells: { A2: 0.2, A3: 10 },
    expected: 0.86466472,
    digits: 8,
    source: "Excel",
    note: "EXPONDIST function, example row 1: cumulative distribution at x=0.2 with lambda=10; Microsoft prints 0.86466472.",
  },
  {
    formula: "EXPONDIST(A2,A3,FALSE)",
    cells: { A2: 0.2, A3: 10 },
    expected: 1.35335283,
    digits: 9,
    source: "Excel",
    note: "EXPONDIST function, example row 2: probability density at x=0.2 with lambda=10; Microsoft prints 1.35335283.",
  },
  {
    formula: "EXPONDIST(0,1,TRUE)",
    expected: 0,
    digits: 15,
    source: "pinned",
    note: "pinned: the exponential CDF is 1-e^(-lambda*x); at x=0 that is 1-e^0 = 1-1 = 0 exactly, for any lambda, same rule as EXPON.DIST (SPEC-127 pin).",
  },
  {
    formula: "EXPONDIST(1,0,TRUE)",
    expected: { error: "#NUM!" },
    source: "Excel",
    note: 'EXPONDIST function, Remarks: "If lambda <= 0, EXPONDIST returns the #NUM! error value." — lambda=0.',
  },
];
