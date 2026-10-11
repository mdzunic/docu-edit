// SPEC-127 "Conformance tables": the Text category's rows. Every row here is
// "pinned" — hand-derived from this engine's documented semantics
// (src/lib/spreadsheet/formula/functions/text.ts's comments, for the already-
// implemented functions) or from SPEC-127's Function pins (Windows-1252
// CHAR/CODE, UNICHAR/UNICODE code points, CLEAN's U+0000-U+001F, FIXED/
// DOLLAR's rounding and en-US format codes, ARRAYTOTEXT's concise/strict
// forms) — never from running this engine or from memory of Microsoft's
// printed examples. No fetched Microsoft example tables are transcribed in
// this file (see the final report for the per-function list).
import type { ConformanceRow } from "./row";

export const TEXT_ROWS: ConformanceRow[] = [
  // CONCAT
  {
    formula: 'CONCAT("a","b","c")',
    expected: "abc",
    source: "pinned",
    note: "pinned: text.ts — CONCAT joins its value arguments with no delimiter.",
  },
  {
    formula: "CONCAT(A1:B2)",
    cells: { A1: "a", B1: "b", A2: "c", B2: "d" },
    expected: "abcd",
    source: "pinned",
    note: "pinned: text.ts — CONCAT reads a range row-major: A1,B1,A2,B2.",
  },
  {
    formula: 'CONCAT("x",1,TRUE)',
    expected: "x1TRUE",
    source: "pinned",
    note: "pinned: numbers and booleans are coerced to text and joined.",
  },
  {
    formula: "CONCAT()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs at least 1) is #VALUE!.",
  },
  // CONCATENATE
  {
    formula: 'CONCATENATE("Hello"," ","World")',
    expected: "Hello World",
    source: "pinned",
    note: "pinned: text.ts — CONCATENATE joins its value arguments with no delimiter.",
  },
  {
    formula: "CONCATENATE(A1,A2)",
    cells: { A1: "Stream", A2: " population" },
    expected: "Stream population",
    source: "pinned",
    note: "pinned: joins the two cell values in order.",
  },
  {
    formula: 'CONCATENATE("a",1,TRUE)',
    expected: "a1TRUE",
    source: "pinned",
    note: "pinned: numbers and booleans are coerced to text and joined.",
  },
  {
    formula: "CONCATENATE()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs at least 1) is #VALUE!.",
  },
  // TEXTJOIN
  {
    formula: 'TEXTJOIN(",",TRUE,"a","b","","c")',
    expected: "a,b,c",
    source: "pinned",
    note: "pinned: text.ts — ignore_empty TRUE skips the empty item.",
  },
  {
    formula: 'TEXTJOIN(",",FALSE,"a","b","","c")',
    expected: "a,b,,c",
    source: "pinned",
    note: "pinned: ignore_empty FALSE keeps the empty item, giving a doubled delimiter.",
  },
  {
    formula: 'TEXTJOIN("-",TRUE,A1:A3)',
    cells: { A1: "x", A2: "", A3: "y" },
    expected: "x-y",
    source: "pinned",
    note: "pinned: reads the range row-major, skipping the empty A2.",
  },
  {
    formula: 'TEXTJOIN(",",TRUE)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (2, needs at least 3) is #VALUE!.",
  },
  // LEN
  {
    formula: 'LEN("Phoenix")',
    expected: 7,
    source: "pinned",
    note: "pinned: text.ts — LEN counts characters.",
  },
  {
    formula: 'LEN("")',
    expected: 0,
    source: "pinned",
    note: "pinned: an empty text has length 0.",
  },
  {
    formula: "LEN(A1)",
    cells: { A1: 12345 },
    expected: 5,
    source: "pinned",
    note: "pinned: a number argument is coerced to text (\"12345\") before measuring.",
  },
  {
    formula: "LEN()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // LEFT
  {
    formula: 'LEFT("Sweden",3)',
    expected: "Swe",
    source: "pinned",
    note: "pinned: text.ts — LEFT takes the first num_chars characters.",
  },
  {
    formula: 'LEFT("Sweden")',
    expected: "S",
    source: "pinned",
    note: "pinned: num_chars defaults to 1.",
  },
  {
    formula: 'LEFT("Sweden",0)',
    expected: "",
    source: "pinned",
    note: "pinned: zero characters gives an empty text.",
  },
  {
    formula: 'LEFT("Sweden",-1)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — a negative num_chars is #VALUE!.",
  },
  // RIGHT
  {
    formula: 'RIGHT("Sweden",3)',
    expected: "den",
    source: "pinned",
    note: "pinned: text.ts — RIGHT takes the last num_chars characters.",
  },
  {
    formula: 'RIGHT("Sweden")',
    expected: "n",
    source: "pinned",
    note: "pinned: num_chars defaults to 1.",
  },
  {
    formula: 'RIGHT("Sweden",0)',
    expected: "",
    source: "pinned",
    note: "pinned: zero characters gives an empty text.",
  },
  {
    formula: 'RIGHT("Sweden",-1)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — a negative num_chars is #VALUE!.",
  },
  // MID
  {
    formula: 'MID("Fluid Flow",1,5)',
    expected: "Fluid",
    source: "pinned",
    note: "pinned: text.ts — characters 1 through 5.",
  },
  {
    formula: 'MID("Fluid Flow",7,20)',
    expected: "Flow",
    source: "pinned",
    note: "pinned: starting at 7 for 20 characters is clipped to the text's actual end, \"Flow\".",
  },
  {
    formula: 'MID("Fluid Flow",20,5)',
    expected: "",
    source: "pinned",
    note: "pinned: a start past the text's end gives an empty text.",
  },
  {
    formula: 'MID("Fluid Flow",0,5)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — a start_num below 1 is #VALUE!.",
  },
  // UPPER
  {
    formula: 'UPPER("total")',
    expected: "TOTAL",
    source: "pinned",
    note: "pinned: text.ts — UPPER converts to upper case.",
  },
  {
    formula: 'UPPER("Yield")',
    expected: "YIELD",
    source: "pinned",
    note: "pinned: already-capitalised letters are also upper-cased.",
  },
  {
    formula: 'UPPER("123abc")',
    expected: "123ABC",
    source: "pinned",
    note: "pinned: digits are unaffected, letters are upper-cased.",
  },
  {
    formula: "UPPER()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // LOWER
  {
    formula: 'LOWER("E. E. Cummings")',
    expected: "e. e. cummings",
    source: "pinned",
    note: "pinned: text.ts — LOWER converts to lower case.",
  },
  {
    formula: 'LOWER("Total")',
    expected: "total",
    source: "pinned",
    note: "pinned: a capitalised word is lower-cased.",
  },
  {
    formula: 'LOWER("123ABC")',
    expected: "123abc",
    source: "pinned",
    note: "pinned: digits are unaffected, letters are lower-cased.",
  },
  {
    formula: "LOWER()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // PROPER
  {
    formula: 'PROPER("this is a title")',
    expected: "This Is A Title",
    source: "pinned",
    note: "pinned: text.ts — every letter following a non-letter (including the first) is capitalised, every other letter lower-cased.",
  },
  {
    formula: 'PROPER("2-way street")',
    expected: "2-Way Street",
    source: "pinned",
    note: 'pinned: the "w" follows the non-letter "-", so it capitalises too: "2-Way Street".',
  },
  {
    formula: 'PROPER("76BudGet")',
    expected: "76Budget",
    source: "pinned",
    note: 'pinned: "B" follows the non-letter "6" and capitalises; the "G" follows the letter "d" and lower-cases: "76Budget".',
  },
  {
    formula: "PROPER()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // TRIM
  {
    formula: 'TRIM("   First Quarter Earnings   ")',
    expected: "First Quarter Earnings",
    source: "pinned",
    note: "pinned: text.ts — TRIM removes leading/trailing spaces.",
  },
  {
    formula: 'TRIM("a    b")',
    expected: "a b",
    source: "pinned",
    note: "pinned: an inner run of spaces collapses to one.",
  },
  {
    formula: 'TRIM("")',
    expected: "",
    source: "pinned",
    note: "pinned: an empty text stays empty.",
  },
  {
    formula: "TRIM()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // SUBSTITUTE
  {
    formula: 'SUBSTITUTE("Quarter 1, 2008","1","2")',
    expected: "Quarter 2, 2008",
    source: "pinned",
    note: 'pinned: text.ts — with no instance_num, every occurrence of "1" is replaced.',
  },
  {
    formula: 'SUBSTITUTE("Sales Data","Sales","Cost")',
    expected: "Cost Data",
    source: "pinned",
    note: "pinned: replaces the one occurrence.",
  },
  {
    formula: 'SUBSTITUTE("aaa","a","b",2)',
    expected: "aba",
    source: "pinned",
    note: "pinned: instance_num 2 replaces only the second \"a\".",
  },
  {
    formula: 'SUBSTITUTE("x","a","b",0)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — an instance_num below 1 is #VALUE!.",
  },
  // REPLACE
  {
    formula: 'REPLACE("abcdefghijk",6,5,"*")',
    expected: "abcde*k",
    source: "pinned",
    note: "pinned: characters 6–10 (\"fghij\") are replaced by \"*\": \"abcde\" + \"*\" + \"k\" (Microsoft's REPLACE example row 2 prints the same).",
  },
  {
    formula: 'REPLACE("2009",3,2,"10")',
    expected: "2010",
    source: "pinned",
    note: 'pinned: characters 3-4 ("09") are replaced by "10": "20" + "10" = "2010".',
  },
  {
    formula: 'REPLACE("Hello",1,0,"X")',
    expected: "XHello",
    source: "pinned",
    note: "pinned: num_chars 0 inserts \"X\" before position 1 without removing anything: \"\" + \"X\" + \"Hello\".",
  },
  {
    formula: 'REPLACE("abc",0,1,"x")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — a start_num below 1 is #VALUE!.",
  },
  // FIND
  {
    formula: 'FIND("M","Miriam McGovern")',
    expected: 1,
    source: "pinned",
    note: 'pinned: text.ts — FIND is case-sensitive; the first "M" is at position 1.',
  },
  {
    formula: 'FIND("m","Miriam McGovern")',
    expected: 6,
    source: "pinned",
    note: 'pinned: case-sensitive; the first lower-case "m" is at position 6 (within "Miriam").',
  },
  {
    formula: 'FIND("Miriam","Miriam McGovern",2)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: 'pinned: starting the search at position 2 skips past the only "Miriam", so it is not found, #VALUE!.',
  },
  {
    formula: 'FIND("z","Miriam McGovern")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — not found is #VALUE!.",
  },
  // SEARCH
  {
    formula: 'SEARCH("m","Miriam McGovern")',
    expected: 1,
    source: "pinned",
    note: 'pinned: text.ts — SEARCH is case-insensitive, so the leading "M" matches "m" at position 1.',
  },
  {
    formula: 'SEARCH("McGovern","Miriam McGovern")',
    expected: 8,
    source: "pinned",
    note: '"McGovern" starts at position 8.',
  },
  {
    formula: 'SEARCH("r*n","Miriam McGovern")',
    expected: 3,
    source: "pinned",
    note: 'pinned: "*" is a wildcard for any run of characters; the earliest "r" that is eventually followed by an "n" is position 3, and the text does end in "n".',
  },
  {
    formula: 'SEARCH("xyz","Miriam McGovern")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — not found is #VALUE!.",
  },
  // TEXT
  {
    formula: 'TEXT(1234.567,"$#,##0.00")',
    expected: "$1,234.57",
    source: "pinned",
    note: "pinned: text.ts formats through formatValue; rounds to the format's two decimals and groups thousands.",
  },
  {
    formula: 'TEXT(0.285,"0.0%")',
    expected: "28.5%",
    source: "pinned",
    note: "pinned: a percentage format multiplies by 100 and appends \"%\".",
  },
  {
    formula: 'TEXT(45366,"mm/dd/yyyy")',
    expected: "03/15/2024",
    source: "pinned",
    note: "pinned: serial 45366 is 2024-03-15, formatted as mm/dd/yyyy.",
  },
  {
    formula: "TEXT(1)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (1, needs 2) is #VALUE!.",
  },
  // VALUE
  {
    formula: 'VALUE("$1,000")',
    expected: 1000,
    source: "pinned",
    note: "pinned: text.ts's valueFromText parses en-US currency text (\"$\" and \",\" grouping) to a number, consistent with the engine's confirmed '\"1\"+\"$4.00\"' behaviour.",
  },
  {
    formula: 'VALUE("16:48:00")',
    expected: 0.7,
    source: "pinned",
    note: "pinned: 16:48:00 is (16*3600+48*60)/86400 = 60480/86400 = 0.7 of a day.",
  },
  {
    formula: "VALUE(A1)",
    cells: { A1: null },
    expected: 0,
    source: "pinned",
    note: "pinned: text.ts — VALUE of an empty cell is 0.",
  },
  {
    formula: 'VALUE("abc")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — text that is not a number, date or time is #VALUE!.",
  },
  // REPT
  {
    formula: 'REPT("-",10)',
    expected: "----------",
    source: "pinned",
    note: "pinned: text.ts — REPT repeats the text number_times.",
  },
  {
    formula: 'REPT("ab",3)',
    expected: "ababab",
    source: "pinned",
    note: "pinned: \"ab\" repeated 3 times.",
  },
  {
    formula: 'REPT("x",0)',
    expected: "",
    source: "pinned",
    note: "pinned: zero repetitions gives an empty text.",
  },
  {
    formula: 'REPT("x",-1)',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text.ts — a negative number_times is #VALUE!.",
  },
  // EXACT
  {
    formula: 'EXACT("Yes","Yes")',
    expected: true,
    source: "pinned",
    note: "pinned: text.ts — identical text, case included, is TRUE.",
  },
  {
    formula: 'EXACT("Yes","yes")',
    expected: false,
    source: "pinned",
    note: "pinned: EXACT is case-sensitive, so differing case is FALSE.",
  },
  {
    formula: 'EXACT("","")',
    expected: true,
    source: "pinned",
    note: "pinned: two empty texts are identical.",
  },
  {
    formula: 'EXACT("a")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (1, needs 2) is #VALUE!.",
  },
  // CHAR
  {
    formula: "CHAR(65)",
    expected: "A",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — CHAR maps through Windows-1252; byte 65 is \"A\" (as in ASCII/Latin-1).",
  },
  {
    formula: "CHAR(97)",
    expected: "a",
    source: "pinned",
    note: "pinned: byte 97 is \"a\".",
  },
  {
    formula: "CHAR(128)",
    expected: "€",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — Windows-1252 byte 0x80 (128) is € (U+20AC), unlike Latin-1 where that byte is a C1 control code.",
  },
  {
    formula: "CHAR()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // CODE
  {
    formula: 'CODE("A")',
    expected: 65,
    source: "pinned",
    note: "pinned: \"A\" is Windows-1252 byte 65.",
  },
  {
    formula: 'CODE("€")',
    expected: 128,
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — € (U+20AC) is Windows-1252 byte 0x80 (128).",
  },
  {
    formula: 'CODE("☃")',
    expected: 63,
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — ☃ (U+2603 SNOWMAN) has no Windows-1252 byte, so CODE of an unmappable character is 63.",
  },
  {
    formula: "CODE(#N/A)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate.",
  },
  // UNICHAR
  {
    formula: "UNICHAR(65)",
    expected: "A",
    source: "pinned",
    note: "pinned: UNICHAR reads a Unicode code point directly; 65 is \"A\".",
  },
  {
    formula: "UNICHAR(128512)",
    expected: "😀",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — 128512 is U+1F600 GRINNING FACE, rendered as a whole surrogate pair.",
  },
  {
    formula: "UNICHAR(0)",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — UNICHAR of 0 is #VALUE!.",
  },
  {
    formula: "UNICHAR(55296)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — 55296 is U+D800, a surrogate code point, so UNICHAR is #N/A.",
  },
  // UNICODE
  {
    formula: 'UNICODE("A")',
    expected: 65,
    source: "pinned",
    note: "pinned: UNICODE reads the Unicode code point of the first character, \"A\" is 65.",
  },
  {
    formula: 'UNICODE("😀")',
    expected: 128512,
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — a surrogate pair is read whole, giving 128512 (U+1F600).",
  },
  {
    formula: 'UNICODE("€")',
    expected: 8364,
    source: "pinned",
    note: "pinned: UNICODE reads the actual Unicode code point (U+20AC = 8364 decimal), unlike CODE which maps through Windows-1252 (128).",
  },
  {
    formula: "UNICODE(#N/A)",
    expected: { error: "#N/A" },
    source: "pinned",
    note: "pinned: SPEC-120 errors propagate.",
  },
  // CLEAN
  {
    formula: 'CLEAN("a"&CHAR(9)&"b")',
    expected: "ab",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — CLEAN removes U+0000-U+001F; CHAR(9) is a tab (U+0009).",
  },
  {
    formula: 'CLEAN(CHAR(10)&"text"&CHAR(13))',
    expected: "text",
    source: "pinned",
    note: "pinned: CHAR(10)/CHAR(13) (line feed U+000A, carriage return U+000D) are both removed.",
  },
  {
    formula: 'CLEAN("clean text")',
    expected: "clean text",
    source: "pinned",
    note: "pinned: text with no control characters is unchanged.",
  },
  {
    formula: "CLEAN()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // FIXED
  {
    formula: "FIXED(1234.567,2)",
    expected: "1,234.57",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — rounds half away from zero to 2 decimals and groups thousands (no_commas defaults to FALSE).",
  },
  {
    formula: "FIXED(1234.567,-2)",
    expected: "1,200",
    source: "pinned",
    note: "pinned: a negative decimals rounds to that power of ten left of the point: 1234.567 to the nearest hundred is 1,200.",
  },
  {
    formula: "FIXED(-1234.567,1,TRUE)",
    expected: "-1234.6",
    source: "pinned",
    note: "pinned: rounds half away from zero to 1 decimal (-1234.6) and no_commas TRUE omits the thousands separator.",
  },
  {
    formula: "FIXED()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs at least 1) is #VALUE!.",
  },
  // DOLLAR
  {
    formula: "DOLLAR(1234.567,2)",
    expected: "$1,234.57",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — en-US DOLLAR format $#,##0.00;($#,##0.00); rounds half away from zero.",
  },
  {
    formula: "DOLLAR(1234.567,-2)",
    expected: "$1,200",
    source: "pinned",
    note: "pinned: a negative decimals rounds to the nearest hundred, 1,200.",
  },
  {
    formula: "DOLLAR(-1234.567,2)",
    expected: "($1,234.57)",
    source: "pinned",
    note: "pinned: SPEC-127's Function pins — en-US wraps a negative amount in parentheses rather than a minus sign.",
  },
  {
    formula: "DOLLAR()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs at least 1) is #VALUE!.",
  },
  // NUMBERVALUE
  {
    formula: 'NUMBERVALUE("2500.27")',
    expected: 2500.27,
    source: "pinned",
    note: "pinned: default en-US separators (\".\" decimal, \",\" group) parse this directly.",
  },
  {
    formula: 'NUMBERVALUE("2.500,27",",",".")',
    expected: 2500.27,
    source: "pinned",
    note: "pinned: explicit decimal_separator \",\" and group_separator \".\" reinterpret the European-style text as the same number, 2500.27.",
  },
  {
    formula: 'NUMBERVALUE("3.5%")',
    expected: 0.035,
    source: "pinned",
    note: "pinned: a trailing \"%\" divides the parsed number by 100: 3.5/100 = 0.035.",
  },
  {
    formula: 'NUMBERVALUE("abc")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: text that cannot be interpreted as a number is #VALUE!.",
  },
  // T
  {
    formula: 'T("Rainfall")',
    expected: "Rainfall",
    source: "pinned",
    note: "pinned: T returns a text value unchanged.",
  },
  {
    formula: "T(19)",
    expected: "",
    source: "pinned",
    note: "pinned: T of a number (not text) is an empty text.",
  },
  {
    formula: "T(TRUE)",
    expected: "",
    source: "pinned",
    note: "pinned: T of a boolean (not text) is an empty text.",
  },
  {
    formula: "T()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs 1) is #VALUE!.",
  },
  // TEXTBEFORE
  {
    formula: 'TEXTBEFORE("Jeff Smith"," ")',
    expected: "Jeff",
    source: "pinned",
    note: "pinned: the text before the first space.",
  },
  {
    formula: 'TEXTBEFORE("Jeff Smith Ward"," ",2)',
    expected: "Jeff Smith",
    source: "pinned",
    note: "pinned: instance_num 2 is the text before the second space.",
  },
  {
    formula: 'TEXTBEFORE("Jeff Smith Ward"," ",-1)',
    expected: "Jeff Smith",
    source: "pinned",
    note: "pinned: a negative instance_num counts from the end; -1 is the text before the last (second) space.",
  },
  {
    formula: 'TEXTBEFORE("Jeff Smith")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (1, needs at least 2) is #VALUE!.",
  },
  // TEXTAFTER
  {
    formula: 'TEXTAFTER("Jeff Smith"," ")',
    expected: "Smith",
    source: "pinned",
    note: "pinned: the text after the first space.",
  },
  {
    formula: 'TEXTAFTER("Jeff Smith Ward"," ",2)',
    expected: "Ward",
    source: "pinned",
    note: "pinned: instance_num 2 is the text after the second space.",
  },
  {
    formula: 'TEXTAFTER("Jeff Smith Ward"," ",-1)',
    expected: "Ward",
    source: "pinned",
    note: "pinned: a negative instance_num counts from the end; -1 is the text after the last space.",
  },
  {
    formula: 'TEXTAFTER("Jeff Smith")',
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (1, needs at least 2) is #VALUE!.",
  },
  // VALUETOTEXT
  {
    formula: 'VALUETOTEXT("Hello")',
    expected: "Hello",
    source: "pinned",
    note: "pinned: format 0 (concise, default) renders text unquoted.",
  },
  {
    formula: 'VALUETOTEXT("Hello",1)',
    expected: '"Hello"',
    source: "pinned",
    note: "pinned: format 1 (strict) quotes a text value.",
  },
  {
    formula: "VALUETOTEXT(123)",
    expected: "123",
    source: "pinned",
    note: "pinned: a number renders the same in concise or strict form.",
  },
  {
    formula: "VALUETOTEXT()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs at least 1) is #VALUE!.",
  },
  // ARRAYTOTEXT
  {
    formula: "ARRAYTOTEXT({1,2,3})",
    expected: "1, 2, 3",
    source: "pinned",
    note: 'pinned: SPEC-127\'s Function pins — concise format (0, default) joins every value with ", ".',
  },
  {
    formula: 'ARRAYTOTEXT({"a",1;"b",2},1)',
    expected: '{"a",1;"b",2}',
    source: "pinned",
    note: 'pinned: SPEC-127\'s Function pins — strict format (1) wraps the result in {…}, uses "," within a row and ";" between rows, and quotes text values.',
  },
  {
    formula: 'ARRAYTOTEXT({"a","b"})',
    expected: "a, b",
    source: "pinned",
    note: 'pinned: concise format joins text values with ", ", unquoted.',
  },
  {
    formula: "ARRAYTOTEXT()",
    expected: { error: "#VALUE!" },
    source: "pinned",
    note: "pinned: SPEC-120 — wrong argument count (0, needs at least 1) is #VALUE!.",
  },

  // SUBSTITUTE … CODE: transcribed from Microsoft's pages (FIND via the en-GB
  // mirror of the same page; SEARCH's A2 re-checked as "Statements").
  { formula: 'SUBSTITUTE(A2,"Sales","Cost")', cells: { A2: "Sales Data" }, expected: "Cost Data", source: "Excel", note: "SUBSTITUTE function, Examples, row 2" },
  { formula: 'SUBSTITUTE(A3,"1","2",1)', cells: { A3: "Quarter 1, 2008" }, expected: "Quarter 2, 2008", source: "Excel", note: "SUBSTITUTE function, Examples, row 3" },
  { formula: 'SUBSTITUTE(A4,"1","2",3)', cells: { A4: "Quarter 1, 2011" }, expected: "Quarter 1, 2012", source: "Excel", note: "SUBSTITUTE function, Examples, row 4" },
  { formula: 'SUBSTITUTE("abc abc abc","abc","X",0)', expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: an instance_num below 1 is #VALUE! (SPEC-120's SUBSTITUTE)" },
  { formula: 'REPLACE(A2,6,5,"*")', cells: { A2: "abcdefghijk" }, expected: "abcde*k", source: "Excel", note: "REPLACE function, Examples, row 2" },
  { formula: 'REPLACE(A3,3,2,"10")', cells: { A3: "2009" }, expected: "2010", source: "Excel", note: "REPLACE function, Examples, row 3" },
  { formula: 'REPLACE(A4,1,3,"@")', cells: { A4: "123456" }, expected: "@456", source: "Excel", note: "REPLACE function, Examples, row 4" },
  { formula: 'REPLACE("abcdef",0,2,"X")', expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: a start_num below 1 is #VALUE! (SPEC-120's REPLACE)" },
  { formula: 'REPLACE(A3,SEARCH(A4,A3),6,"Amount")', cells: { A3: "Profit Margin", A4: "margin" }, expected: "Profit Amount", source: "Excel", note: "SEARCH function, Examples, row 3" },
  { formula: 'FIND("M","Miriam McGovern")', expected: 1, source: "Excel", note: "FIND function, Examples, row 1" },
  { formula: 'FIND("m","Miriam McGovern")', expected: 6, source: "Excel", note: "FIND function, Examples, row 2" },
  { formula: 'FIND("M","Miriam McGovern",3)', expected: 8, source: "Excel", note: "FIND function, Examples, row 3" },
  { formula: 'FIND("z","Miriam McGovern")', expected: { error: "#VALUE!" }, source: "Excel", note: "FIND function, Remarks: find_text not in within_text is #VALUE!" },
  { formula: 'FIND("M","Miriam McGovern",0)', expected: { error: "#VALUE!" }, source: "Excel", note: "FIND function, Remarks: start_num not greater than zero is #VALUE!" },
  { formula: 'FIND("M","Miriam McGovern",16)', expected: { error: "#VALUE!" }, source: "Excel", note: "FIND function, Remarks: start_num past within_text's length (15) is #VALUE!" },
  { formula: 'SEARCH("e",A2,6)', cells: { A2: "Statements" }, expected: 7, source: "Excel", note: "SEARCH function, Examples, row 1" },
  { formula: "SEARCH(A4,A3)", cells: { A3: "Profit Margin", A4: "margin" }, expected: 8, source: "Excel", note: "SEARCH function, Examples, row 2" },
  { formula: 'SEARCH("""",A5)', cells: { A5: 'The "boss" is here.' }, expected: 5, source: "Excel", note: "SEARCH function, Examples, row 5" },
  { formula: 'SEARCH("xyz","Profit Margin")', expected: { error: "#VALUE!" }, source: "Excel", note: "SEARCH function, Remarks: find_text not found is #VALUE!" },
  { formula: 'SEARCH("M?rgin","Profit Margin")', expected: 8, source: "pinned", note: "pinned: ? matches one character, so M?rgin matches Margin at 8" },
  { formula: 'MID(A3,SEARCH(" ",A3)+1,4)', cells: { A3: "Profit Margin" }, expected: "Marg", source: "Excel", note: "SEARCH function, Examples, row 4" },
  { formula: 'TEXT(1234.567,"$#,##0.00")', expected: "$1,234.57", source: "Excel", note: "TEXT function, Examples, row 1" },
  { formula: 'TEXT(0.285,"0.0%")', expected: "28.5%", source: "Excel", note: "TEXT function, Examples, row 5" },
  { formula: 'TEXT(1234,"0000000")', expected: "0001234", source: "Excel", note: "TEXT function, Examples, row 8" },
  { formula: "TEXT(A1,\"0\")", cells: { A1: { error: "#N/A" } }, expected: { error: "#N/A" }, source: "pinned", note: "pinned: a value argument's error is returned (SPEC-120 Errors)" },
  { formula: 'VALUE("$1,000")', expected: 1000, source: "Excel", note: "VALUE function, Examples, row 1" },
  { formula: "VALUE(42)", expected: 42, source: "pinned", note: "pinned: a number passes through VALUE (SPEC-120)" },
  { formula: 'VALUE("abc")', expected: { error: "#VALUE!" }, source: "Excel", note: "VALUE function, Remarks: text in no number, date or time format is #VALUE!" },
  { formula: 'REPT("*-",3)', expected: "*-*-*-", source: "Excel", note: "REPT function, Examples, row 1" },
  { formula: 'REPT("-",10)', expected: "----------", source: "Excel", note: "REPT function, Examples, row 2" },
  { formula: 'REPT("x",0)', expected: "", source: "Excel", note: "REPT function, Remarks: number_times 0 returns empty text" },
  { formula: 'REPT("a",-1)', expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: a negative number_times is #VALUE! (SPEC-120's REPT)" },
  { formula: "EXACT(A2,B2)", cells: { A2: "word", B2: "word" }, expected: true, source: "Excel", note: "EXACT function, Examples, row 2" },
  { formula: "EXACT(A3,B3)", cells: { A3: "Word", B3: "word" }, expected: false, source: "Excel", note: "EXACT function, Examples, row 3" },
  { formula: "EXACT(A4,B4)", cells: { A4: "w ord", B4: "word" }, expected: false, source: "Excel", note: "EXACT function, Examples, row 4" },
  { formula: 'EXACT("a")', expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: a wrong argument count is #VALUE! (SPEC-120 Errors)" },
  { formula: "CHAR(65)", expected: "A", source: "Excel", note: "CHAR function, Examples, row 1" },
  { formula: "CHAR(33)", expected: "!", source: "Excel", note: "CHAR function, Examples, row 2" },
  { formula: "CHAR(128)", expected: "€", source: "pinned", note: "pinned: Windows-1252 byte 0x80 is € (U+20AC), SPEC-127's CHAR pin" },
  { formula: "CHAR(65,66)", expected: { error: "#VALUE!" }, source: "pinned", note: "pinned: a wrong argument count is #VALUE! (SPEC-120 Errors)" },
  { formula: 'CODE("A")', expected: 65, source: "Excel", note: "CODE function, Examples, row 1" },
  { formula: 'CODE("!")', expected: 33, source: "Excel", note: "CODE function, Examples, row 2" },
  { formula: 'CODE("€")', expected: 128, source: "pinned", note: "pinned: € is Windows-1252 byte 0x80 (SPEC-127's CODE pin)" },
  { formula: 'CODE("☃")', expected: 63, source: "pinned", note: "pinned: U+2603 has no Windows-1252 byte, so CODE is 63 (SPEC-127's pin)" },
  { formula: "CODE(A1)", cells: { A1: { error: "#N/A" } }, expected: { error: "#N/A" }, source: "pinned", note: "pinned: a value argument's error is returned (SPEC-120 Errors)" },
];
