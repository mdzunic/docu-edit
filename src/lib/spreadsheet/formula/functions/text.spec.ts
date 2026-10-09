// SPEC-120 "The core functions" (text.ts): CONCAT CONCATENATE TEXTJOIN LEN
// LEFT RIGHT MID UPPER LOWER PROPER TRIM SUBSTITUTE REPLACE FIND SEARCH TEXT
// VALUE REPT EXACT. Lengths and positions count UTF-16 code units
// (LEN("😀") is 2); PROPER capitalises every letter after a non-letter
// ("It'S"); TRIM removes U+0020 only and collapses its runs; FIND is
// case-sensitive without wildcards, SEARCH case-insensitive with them, and
// not found is #VALUE!; CONCAT and TEXTJOIN read ranges row-major while
// CONCATENATE takes values; TEXTJOIN's ignore_empty both ways; TEXT formats
// through SPEC-118's formatter — the code read as invariant en-US, the output
// localised — and an invalid code is #VALUE!; VALUE goes through
// valueFromText ("" is #VALUE!, an empty cell 0, "1,5" 1.5 in de-DE); a text
// result longer than 32,767 characters is #VALUE!.
//
// Sources ("Where expected values come from"): "docs" rows are Microsoft's
// function-reference examples, recomputed by hand from formula and data;
// "pinned" rows follow from the rule SPEC-120 states.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../../address";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellError, CellValue, ErrorCode } from "../../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../../numberFormat/locale";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

const err = (code: ErrorCode): CellError => ({ error: code });

/** 2026-10-08 09:30 local, a Thursday — the clock TODAY() and NOW() read. */
const NOW = new Date(2026, 9, 8, 9, 30);

interface Where {
  host?: string;
  /** The host is an array formula's anchor over this range. */
  array?: string;
  locale?: SpreadsheetLocaleTag;
  format?: "xlsx" | "ods";
  dateSystem?: "1900" | "1904";
}

/** What `formula`, in `host` (default Z100) of Sheet1 beside `cells`,
 * stores under `locale` (default en-US) with the clock at NOW — an xlsx
 * workbook in the 1900 system unless `where` says otherwise. */
function value(formula: string, cells: Cells = {}, where: Where = {}) {
  const host = where.host ?? "Z100";
  const self: CellSpec = where.array === undefined ? { formula } : { formula, array: where.array };
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: self } }, { format: where.format, dateSystem: where.dateSystem });
  const at = parseA1(host)!;
  const settings = { locale: SPREADSHEET_LOCALES[where.locale ?? "en-US"] };
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, { ...storedValueContext(model, settings), now: () => NOW });
}

const table = (rows: [formula: string, expected: CellValue][]) => rows;

describe("CONCATENATE, CONCAT and TEXTJOIN", () => {
  it("docs: CONCATENATE joins its values", () => {
    const cells: Cells = { A2: "brook trout", B2: "Andreas", C2: "Hauser", A3: "species", B3: "Fourth", C3: "Pine", A4: 32 };
    expect(value('CONCATENATE("Stream population for ",A2," ",A3," is ",A4,"/mile.")', cells)).toBe(
      "Stream population for brook trout species is 32/mile.",
    );
    expect(value('CONCATENATE(B2," ",C2)', cells)).toBe("Andreas Hauser");
    expect(value('CONCATENATE(C2,", ",B2)', cells)).toBe("Hauser, Andreas");
    expect(value('CONCATENATE(B3," & ",C3)', cells)).toBe("Fourth & Pine");
    expect(value('B3&" & "&C3', cells)).toBe("Fourth & Pine");
  });

  it("CONCATENATE takes value arguments: a range is intersected with the host's row", () => {
    const cells: Cells = { A1: "a1", A2: "a2", A3: "a3" };
    expect(value("CONCATENATE(A1:A3)", cells, { host: "B2" })).toBe("a2");
    expect(value("CONCATENATE(A1:A3)", cells, { host: "B9" })).toEqual(err("#VALUE!"));
  });

  it("CONCAT reads ranges and arrays row-major, an empty cell adding nothing", () => {
    expect(value('CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")')).toBe("The sun will come up tomorrow."); // docs
    expect(value("CONCAT(A1:B2)", { A1: "a1", B1: "b1", A2: "a2", B2: "b2" })).toBe("a1b1a2b2");
    expect(value('CONCAT({"a","b";"c","d"})')).toBe("abcd");
    expect(value("CONCAT(A1:A3)", { A1: "a", A3: "c" })).toBe("ac");
    expect(value("CONCAT(A1:A2,1/4,TRUE)", { A1: "x", A2: 2 })).toBe("x20.25TRUE");
    expect(value('CONCAT("a",A1:A2)', { A1: "b", A2: err("#N/A") })).toEqual(err("#N/A"));
  });

  it("docs: TEXTJOIN joins a column with a delimiter", () => {
    const currencies = ["US Dollar", "Australian Dollar", "Chinese Yuan", "Hong Kong Dollar", "Israeli Shekel", "South Korean Won", "Russian Ruble"];
    const cells: Cells = Object.fromEntries(currencies.map((name, i) => [`A${i + 2}`, name]));
    expect(value('TEXTJOIN(", ",TRUE,A2:A8)', cells)).toBe(currencies.join(", "));
  });

  it("docs: TEXTJOIN reads A2:B8 row-major and ignore_empty decides the empty row", () => {
    const cells: Cells = {};
    for (const [row, n] of [[2, 1], [3, 2], [5, 4], [6, 5], [7, 6], [8, 7]]) {
      cells[`A${row}`] = `a${n}`;
      cells[`B${row}`] = `b${n}`;
    }
    expect(value('TEXTJOIN(", ",TRUE,A2:B8)', cells)).toBe("a1, b1, a2, b2, a4, b4, a5, b5, a6, b6, a7, b7");
    expect(value('TEXTJOIN(", ",FALSE,A2:B8)', cells)).toBe("a1, b1, a2, b2, , , a4, b4, a5, b5, a6, b6, a7, b7");
  });

  it("docs: TEXTJOIN cycles through a range of delimiters", () => {
    const cells: Cells = {
      A1: "City", B1: "State", C1: "Postal Code", D1: "Country",
      A2: "Tulsa", B2: "OK", C2: 74133, D2: "US",
      A3: "Seattle", B3: "WA", C3: 98109, D3: "US",
      A4: "Iselin", B4: "NJ", C4: "08830", D4: "US", // text, keeping its 0
      A5: "Fort Lauderdale", B5: "FL", C5: 33309, D5: "US",
      A6: "Tempe", B6: "AZ", C6: 85285, D6: "US",
      A7: "end",
      A8: ",", B8: ",", C8: ",", D8: ";",
    };
    expect(value("TEXTJOIN(A8:D8,TRUE,A2:D7)", cells)).toBe(
      "Tulsa,OK,74133,US;Seattle,WA,98109,US;Iselin,NJ,08830,US;Fort Lauderdale,FL,33309,US;Tempe,AZ,85285,US;end",
    );
  });

  it.each(
    table([
      ['TEXTJOIN("-",TRUE,"a","","c")', "a-c"],
      ['TEXTJOIN("-",FALSE,"a","","c")', "a--c"],
      ['TEXTJOIN("-",TRUE,A1:A3)', "a"],
      ['TEXTJOIN("-",FALSE,A1:A3)', "a--"],
      ['TEXTJOIN(", ",TRUE,{"a","b";"c","d"})', "a, b, c, d"],
      ['TEXTJOIN(",","x","a")', err("#VALUE!")],
    ]),
  )("=%s is %j (A1 'a', A2:A3 empty)", (formula, expected) => {
    expect(value(formula, { A1: "a" })).toEqual(expected);
  });
});

describe("LEN, LEFT, RIGHT and MID", () => {
  it.each(
    table([
      ['LEN("Phoenix, AZ")', 11], // docs
      ["LEN(A1)", 0], // docs: an empty cell
      ['LEN("     One   ")', 11], // docs
      ['LEN("😀")', 2], // pinned: UTF-16 code units
      ["LEN(1/3)", 17], // "0.333333333333333"
      ["LEN(TRUE)", 4],
      ['LEFT("Sale Price",4)', "Sale"], // docs
      ['LEFT("Sweden")', "S"], // docs: num_chars defaults to 1
      ['LEFT("abc",10)', "abc"], // docs: past the end → all of text
      ['LEFT("abc",0)', ""],
      ['LEFT("abc",-1)', err("#VALUE!")], // docs: num_chars must be ≥ 0
      ['RIGHT("Sale Price",5)', "Price"], // docs
      ['RIGHT("Stock Number")', "r"], // docs
      ['RIGHT("abc",-1)', err("#VALUE!")],
      ['MID("Fluid Flow",1,5)', "Fluid"], // docs
      ['MID("Fluid Flow",7,20)', "Flow"], // docs
      ['MID("Fluid Flow",20,5)', ""], // docs
      ['MID("abc",0,1)', err("#VALUE!")], // docs: start_num < 1
      ['MID("abc",1,-1)', err("#VALUE!")], // docs: negative num_chars
      ['MID("a😀b",2,2)', "😀"],
      ['RIGHT("a😀",2)', "😀"],
      ['LEFT("😀b",2)', "😀"],
    ]),
  )("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("in an array formula", () => {
  it("LEN applies element-wise, an empty cell counting 0", () => {
    const cells: Cells = { A1: "ab", A2: "cde" };
    expect(value("LEN(A1:A3)", cells, { host: "B1", array: "B1:B3" })).toEqual({ kind: "array", rows: 3, cols: 1, data: [2, 3, 0] });
    expect(value("SUM(LEN(A1:A3))", cells, { host: "C1", array: "C1:C1" })).toEqual({ kind: "array", rows: 1, cols: 1, data: [5] });
  });
});

describe("UPPER, LOWER, PROPER and TRIM", () => {
  it.each(
    table([
      ['UPPER("total")', "TOTAL"], // docs
      ['UPPER("Yield")', "YIELD"], // docs
      ['UPPER("café")', "CAFÉ"],
      ['LOWER("E. E. Cummings")', "e. e. cummings"], // docs
      ['LOWER("Apt. 2B")', "apt. 2b"], // docs
      ['PROPER("this is a TITLE")', "This Is A Title"], // docs
      ['PROPER("2-way street")', "2-Way Street"], // docs
      ['PROPER("76BudGet")', "76Budget"], // docs
      ['PROPER("it\'s")', "It'S"], // conformance row (pinned)
      ['TRIM(" First Quarter Earnings ")', "First Quarter Earnings"], // docs
      ['TRIM("  a   b  ")', "a b"], // conformance row (Excel)
      ['TRIM("\u00a0a\u00a0")', "\u00a0a\u00a0"], // docs: U+00A0 is not removed
      ['TRIM(" \ta  b ")', "\ta b"], // nor is a tab
    ]),
  )("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("SUBSTITUTE and REPLACE", () => {
  it.each(
    table([
      ['SUBSTITUTE("Sales Data","Sales","Cost")', "Cost Data"], // docs
      ['SUBSTITUTE("Quarter 1, 2008","1","2",1)', "Quarter 2, 2008"], // docs
      ['SUBSTITUTE("Quarter 1, 2011","1","2",3)', "Quarter 1, 2012"], // docs
      ['SUBSTITUTE("aAa","a","x")', "xAx"], // case-sensitive
      ['SUBSTITUTE("aaaa","aa","b")', "bb"],
      ['SUBSTITUTE("aaa","a","b",4)', "aaa"], // no fourth instance
      ['SUBSTITUTE("aaa","a","b",0)', err("#VALUE!")],
      ['REPLACE("abcdefghijk",6,5,"*")', "abcde*k"], // docs
      ['REPLACE(A3,3,2,"10")', "2010"], // docs: A3 the number 2009
      ['REPLACE(A4,1,3,"@")', "@456"], // docs: A4 the number 123456
      ['REPLACE("abc",0,1,"x")', err("#VALUE!")],
      ['REPLACE("abc",1,-1,"x")', err("#VALUE!")],
    ]),
  )("=%s is %j", (formula, expected) => {
    expect(value(formula, { A3: 2009, A4: 123456 })).toEqual(expected);
  });
});

describe("FIND and SEARCH", () => {
  it.each(
    table([
      ['FIND("M",A2)', 1], // docs: A2 "Miriam McGovern"
      ['FIND("m",A2)', 6], // docs
      ['FIND("M",A2,3)', 8], // docs
      ['MID(A3,1,FIND(" #",A3,1)-1)', "Ceramic Insulators"], // docs
      ['MID(A4,1,FIND(" #",A4,1)-1)', "Copper Coils"], // docs
      ['MID(A5,1,FIND(" #",A5,1)-1)', "Variable Resistors"], // docs
      ['FIND("","abc")', 1], // docs: "" matches at start_num
      ['FIND("","abc",2)', 2],
      ['FIND("*","a*b")', 2], // no wildcards: a literal *
      ['FIND("?","abc")', err("#VALUE!")],
      ['FIND("A","abc")', err("#VALUE!")], // case-sensitive: not found
      ['FIND("a","abc",0)', err("#VALUE!")], // docs: start_num not > 0
      ['FIND("a","abc",4)', err("#VALUE!")], // docs: start_num past the text
    ]),
  )("=%s is %j", (formula, expected) => {
    const cells: Cells = { A2: "Miriam McGovern", A3: "Ceramic Insulators #124-TD45-87", A4: "Copper Coils #12-671-6772", A5: "Variable Resistors #116010" };
    expect(value(formula, cells)).toEqual(expected);
  });

  it.each(
    table([
      ['SEARCH("e",A2,6)', 7], // docs
      ["SEARCH(A4,A3)", 8], // docs
      ['REPLACE(A3,SEARCH(A4,A3),6,"Amount")', "Profit Amount"], // docs
      ['MID(A3,SEARCH(" ",A3)+1,4)', "Marg"], // docs
      ['SEARCH("""",A5)', 5], // docs
      ['MID(A5,SEARCH("""",A5)+1,SEARCH("""",A5,SEARCH("""",A5)+1)-SEARCH("""",A5)-1)', "boss"], // docs
      ['SEARCH("c*b","abcb")', 3], // pinned
      ['SEARCH("B","abc")', 2], // case-insensitive
      ['SEARCH("?c","abc")', 2],
      ['SEARCH("~*","a*b")', 2], // docs: ~ before a literal * or ?
      ['SEARCH("~?","a?b")', 2],
      ['SEARCH("z","abc")', err("#VALUE!")], // docs: not found
      ['SEARCH("a","abc",0)', err("#VALUE!")], // docs: start_num not > 0
      ['SEARCH("a","abc",4)', err("#VALUE!")], // docs: start_num past the text
    ]),
  )("=%s is %j", (formula, expected) => {
    const cells: Cells = { A2: "Statements", A3: "Profit Margin", A4: "margin", A5: 'The "boss" is here.' };
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("TEXT", () => {
  it("reads the code as invariant en-US and localises the output", () => {
    expect(value('TEXT(1234.5,"#,##0.00")')).toBe("1,234.50");
    expect(value('TEXT(1234.5,"#,##0.00")', {}, { locale: "de-DE" })).toBe("1.234,50");
    expect(value('TEXT(DATE(2026,10,8),"dddd")', {}, { locale: "de-DE" })).toBe("Donnerstag");
  });

  it("renders a serial in the workbook's own calendar", () => {
    expect(value('TEXT(0,"yyyy-mm-dd")')).toBe("1900-01-00"); // Excel's day 0
    expect(value('TEXT(60,"d-mmm-yyyy")')).toBe("29-Feb-1900");
    expect(value('TEXT(0,"yyyy-mm-dd")', {}, { dateSystem: "1904" })).toBe("1904-01-01");
    expect(value('TEXT(0,"yyyy-mm-dd")', {}, { format: "ods" })).toBe("1899-12-31"); // SPEC-118's proleptic calendar
  });

  it.each(
    table([
      ['TEXT(0.5,"0%")', "50%"], // conformance row (Excel)
      ['TEXT(1234.567,"$#,##0.00")', "$1,234.57"], // docs
      ['TEXT(0.285,"0.0%")', "28.5%"], // docs
      ['TEXT(4.34,"# ?/?")', "4 1/3"], // docs
      ['TRIM(TEXT(0.34,"# ?/?"))', "1/3"], // docs
      ['TEXT(12200000,"0.00E+00")', "1.22E+07"], // docs
      ['TEXT(1234,"0000000")', "0001234"], // docs
      ['TEXT(TODAY(),"MM/DD/YY")', "10/08/26"], // docs' code, the injected clock
      ['TEXT(TODAY(),"DDDD")', "Thursday"],
      ['TEXT(NOW(),"H:MM AM/PM")', "9:30 AM"],
      ['TEXT("abc","0.00")', "abc"], // text passes through a number format
      ['TEXT(A1,"0.00")', "0.00"], // an empty cell is 0
      ['TEXT(1/0,"0")', err("#DIV/0!")],
      ['TEXT(1,"[Red")', err("#VALUE!")], // pinned: an unterminated bracket
      ['TEXT(1,"0.0""")', err("#VALUE!")], // pinned: an unterminated quote
    ]),
  )("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("VALUE", () => {
  it("docs: currency and time text", () => {
    expect(value('VALUE("$1,000")')).toBe(1000);
    expect(value('VALUE("16:48:00")-VALUE("12:00:00")')).toBeCloseTo(0.2, 12);
  });

  it.each<[formula: string, cells: Cells, locale: SpreadsheetLocaleTag, expected: CellValue]>([
    ['VALUE("")', {}, "en-US", err("#VALUE!")],
    ["VALUE(A1)", {}, "en-US", 0], // an empty cell
    ["VALUE(5)", {}, "en-US", 5], // a number passes through
    ["VALUE(A1)", { A1: 5 }, "en-US", 5],
    ['VALUE("abc")', {}, "en-US", err("#VALUE!")],
    ['VALUE("TRUE")', {}, "en-US", err("#VALUE!")], // booleans are never parsed from text
    ['VALUE("1,5")', {}, "en-US", err("#VALUE!")], // pinned
    ['VALUE("1,5")', {}, "de-DE", 1.5], // pinned
    ['VALUE("1,000")', {}, "en-US", 1000],
    ['VALUE("1.234,5")', {}, "de-DE", 1234.5],
    ['VALUE("2026-10-08")', {}, "en-US", 46303], // pinned: ISO in every locale
    ['VALUE("2026-10-08")', {}, "de-DE", 46303],
    ['VALUE("10%")', {}, "en-US", 0.1],
    ['VALUE("(5)")', {}, "en-US", -5],
    ['VALUE(" 12 ")', {}, "en-US", 12], // U+0020 around the number is trimmed
  ])("=%s over %j in %s is %j", (formula, cells, locale, expected) => {
    expect(value(formula, cells, { locale })).toEqual(expected);
  });
});

describe("REPT and EXACT", () => {
  it.each(
    table([
      ['REPT("*-",3)', "*-*-*-"], // docs
      ['REPT("-",10)', "----------"], // docs
      ['REPT("ab",0)', ""], // docs: 0 → empty text
      ['REPT("ab",2.9)', "abab"], // docs: truncated
      ['REPT("a",-1)', err("#VALUE!")],
      ['EXACT("word","word")', true], // docs
      ['EXACT("Word","word")', false], // docs
      ['EXACT("w ord","word")', false], // docs
      ['EXACT(1,"1")', true],
    ]),
  )("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("text results longer than 32,767 characters", () => {
  it("are fine up to the limit", () => {
    expect(value('LEN(REPT("a",32767))')).toBe(32767);
  });

  it.each([
    'REPT("a",40000)', // docs: REPT cannot exceed 32,767 characters
    'CONCATENATE(REPT("a",32767),"b")',
    'CONCAT(REPT("a",20000),REPT("b",20000))', // docs
    'TEXTJOIN(",",TRUE,REPT("a",20000),REPT("b",20000))', // docs
    'SUBSTITUTE(REPT("a",20000),"a","aa")',
    'REPLACE(REPT("a",32767),1,0,"b")',
    'REPT("a",32767)&"b"',
  ])("=%s is #VALUE!", (formula) => {
    expect(value(formula)).toEqual(err("#VALUE!"));
  });
});
