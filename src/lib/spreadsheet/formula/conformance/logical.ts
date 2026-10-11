// SPEC-127 "Conformance tables": the Logical conformance rows.
// Sources: Microsoft's function reference pages
// (support.microsoft.com/en-us/excel/functions/<name>-function) and their
// Remarks, and SPEC-120/127's stated rules (spec.md "Function pins",
// "Errors") for pinned rows. IF, TRUE, FALSE and IFNA's pages show their
// examples only as screenshots with no recoverable cell data (confirmed by
// repeated fetch), so those four are pinned from the functions' own stated
// argument-default and zero-argument rules. Never this engine's own output.
import type { ConformanceRow } from "./row";

export const LOGICAL_ROWS: ConformanceRow[] = [
  // --- IF --------------------------------------------------------------
  // IF function: page has no transcribable example table (screenshots
  // only); rows pinned from the stated syntax (value_if_true/value_if_false
  // default and optional).
  { formula: "IF(TRUE,1,2)", expected: 1, source: "pinned", note: "pinned: IF function — logical_test TRUE returns value_if_true." },
  { formula: "IF(FALSE,1,2)", expected: 2, source: "pinned", note: "pinned: IF function — logical_test FALSE returns value_if_false." },
  { formula: 'IF(1=1,"yes","no")', expected: "yes", source: "pinned", note: "pinned: IF function — an expression logical_test." },
  { formula: "IF(TRUE)", expected: true, source: "pinned", note: 'pinned: IF function syntax — value_if_true omitted defaults to TRUE.' },
  { formula: "IF(FALSE)", expected: false, source: "pinned", note: "pinned: IF function syntax — value_if_false omitted defaults to FALSE." },
  {
    formula: "IF(1/0,1,2)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: SPEC-120 Errors — an error argument (logical_test here) propagates; IF is not lazy on its own test.",
  },

  // --- IFS ---------------------------------------------------------------
  {
    formula: 'IFS(A2>89,"A",A2>79,"B",A2>69,"C",A2>59,"D",TRUE,"F")',
    cells: { A2: 95 },
    expected: "A",
    source: "Excel",
    note: 'IFS function, Example 1 (A2=95): first TRUE condition (A2>89) wins.',
  },
  {
    formula: 'IFS(FALSE,"a",FALSE,"b")',
    expected: { error: "#N/A" },
    source: "Excel",
    note: 'IFS function Remarks: "If no TRUE conditions are found, this function returns #N/A error."',
  },
  { formula: 'IFS(TRUE,"only")', expected: "only", source: "pinned", note: "pinned: IFS function — the minimal one-pair call." },
  {
    formula: 'IFS("x","a")',
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'IFS function Remarks: "If a logical_test argument ... resolves to a value other than TRUE or FALSE, this function returns a #VALUE! error."',
  },

  // --- IFERROR -------------------------------------------------------------
  {
    formula: 'IFERROR(A2/B2,"Error in calculation")',
    cells: { A2: 210, B2: 35 },
    expected: 6,
    source: "Excel",
    note: "IFERROR function, Example 1 row 1: 210/35, no error.",
  },
  {
    formula: 'IFERROR(A3/B3,"Error in calculation")',
    cells: { A3: 55, B3: 0 },
    expected: "Error in calculation",
    source: "Excel",
    note: "IFERROR function, Example 1 row 2: 55/0 is #DIV/0!, caught.",
  },
  {
    formula: 'IFERROR(A4/B4,"Error in calculation")',
    cells: { B4: 23 },
    expected: 0,
    source: "Excel",
    note: 'IFERROR function, Example 1 row 3: A4 empty reads as 0, 0/23 = 0, no error.',
  },
  {
    formula: "IFERROR(1/0,2/0)",
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: IFERROR returns value_if_error unconditionally once value errors, even when value_if_error is itself an error (1/0 caught, then 2/0 evaluated and returned).",
  },

  // --- IFNA ----------------------------------------------------------------
  // IFNA function: page's only example is a VLOOKUP screenshot with no
  // recoverable cell data; rows pinned from the stated value/value_if_na
  // definition (catches #N/A only, unlike IFERROR).
  { formula: 'IFNA(NA(),"missing")', expected: "missing", source: "pinned", note: "pinned: IFNA function — value is #N/A, returns value_if_na." },
  {
    formula: 'IFNA(1/0,"missing")',
    expected: { error: "#DIV/0!" },
    source: "pinned",
    note: "pinned: IFNA function — only #N/A is caught; #DIV/0! propagates unchanged.",
  },
  { formula: 'IFNA(5,"missing")', expected: 5, source: "pinned", note: "pinned: IFNA function — value is not an error, returned as-is." },

  // --- AND -------------------------------------------------------------
  {
    formula: "AND(A1:A2)",
    cells: { A1: "x", A2: "y" },
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'AND function Remarks: "If the specified range contains no logical values, the AND function returns the #VALUE! error."',
  },
  { formula: "AND(TRUE,TRUE)", expected: true, source: "pinned", note: "pinned: AND function — every argument TRUE." },
  { formula: "AND(TRUE,FALSE)", expected: false, source: "pinned", note: "pinned: AND function — one argument FALSE." },
  { formula: "AND(1,1)", expected: true, source: "pinned", note: "pinned: AND function — direct numeric arguments coerce logically, nonzero is TRUE." },

  // --- OR ----------------------------------------------------------------
  {
    formula: "OR(A1:A2)",
    cells: { A1: "x", A2: "y" },
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'OR function Remarks: "If the specified range contains no logical values, OR returns the #VALUE! error value."',
  },
  { formula: "OR(TRUE,FALSE)", expected: true, source: "pinned", note: "pinned: OR function — one argument TRUE." },
  { formula: "OR(FALSE,FALSE)", expected: false, source: "pinned", note: "pinned: OR function — every argument FALSE." },

  // --- NOT -----------------------------------------------------------------
  {
    formula: "NOT(A2>100)",
    cells: { A2: 50 },
    expected: true,
    source: "Excel",
    note: "NOT function example: A2=50, 50 is NOT greater than 100.",
  },
  { formula: "NOT(TRUE)", expected: false, source: "pinned", note: 'pinned: NOT function Remarks — "if logical is TRUE, NOT returns FALSE."' },
  { formula: "NOT(0)", expected: true, source: "pinned", note: "pinned: NOT function — 0 coerces to FALSE, NOT(FALSE) is TRUE." },
  { formula: "NOT(1,2)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-120 Errors — a wrong argument count (max 1) is #VALUE!." },

  // --- XOR -----------------------------------------------------------------
  {
    formula: "XOR(3>0,2<9)",
    expected: false,
    source: "Excel",
    note: "XOR function example: both tests TRUE, an even count of TRUE is FALSE.",
  },
  {
    formula: "XOR(3>12,4>6)",
    expected: false,
    source: "Excel",
    note: "XOR function example: both tests FALSE, an even (zero) count of TRUE is FALSE.",
  },
  { formula: "XOR(TRUE,FALSE)", expected: true, source: "pinned", note: "pinned: XOR function — exactly one TRUE, an odd count, is TRUE." },
  {
    formula: "XOR(A1:A2)",
    cells: { A1: "x", A2: "y" },
    expected: { error: "#VALUE!" },
    source: "Excel",
    note: 'XOR function Remarks: "If the specified range contains no logical values, XOR returns the #VALUE! error value."',
  },

  // --- TRUE ----------------------------------------------------------------
  // TRUE function: page has no transcribable example table (prose only);
  // rows pinned from the zero-argument definition and SPEC-127's explicit
  // wrong-arg-count pin.
  { formula: "TRUE()", expected: true, source: "pinned", note: 'pinned: TRUE function — "Returns the logical value TRUE," no arguments.' },
  { formula: "TRUE(1)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-127 — TRUE() takes no arguments; TRUE(1) is #VALUE!." },
  { formula: "TRUE(A1)", cells: { A1: 5 }, expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-127 — TRUE() takes no arguments; one argument is #VALUE!." },

  // --- FALSE ---------------------------------------------------------------
  { formula: "FALSE()", expected: false, source: "pinned", note: 'pinned: FALSE function — "Returns the logical value FALSE," no arguments.' },
  { formula: "FALSE(1)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-127 — FALSE() takes no arguments; FALSE(1) is #VALUE!., analogous to TRUE(1)" },
  { formula: "FALSE(A1)", cells: { A1: 5 }, expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: SPEC-127 — FALSE() takes no arguments; one argument is #VALUE!." },

  // --- SWITCH --------------------------------------------------------------
  {
    formula: 'SWITCH(WEEKDAY(A2),1,"Sunday",2,"Monday",3,"Tuesday","No match")',
    cells: { A2: 2 },
    expected: "Monday",
    source: "Excel",
    note: "SWITCH function example: A2=2, WEEKDAY(2)=2 maps to the value1=2 pair, result Monday.",
  },
  {
    formula: 'SWITCH(A3,1,"Sunday",2,"Monday",3,"Tuesday")',
    cells: { A3: 99 },
    expected: { error: "#N/A" },
    source: "Excel",
    note: "SWITCH function example: no match and no default (odd final count) returns #N/A.",
  },
  {
    formula: 'SWITCH(A4,1,"Sunday",2,"Monday",3,"Tuesday","No match")',
    cells: { A4: 99 },
    expected: "No match",
    source: "Excel",
    note: "SWITCH function example: no match, trailing unpaired argument is the default.",
  },
  {
    formula: 'SWITCH(A5,1,"Sunday",7,"Saturday","weekday")',
    cells: { A5: 2 },
    expected: "weekday",
    source: "Excel",
    note: "SWITCH function example: non-consecutive match values fall through to the default.",
  },
];
