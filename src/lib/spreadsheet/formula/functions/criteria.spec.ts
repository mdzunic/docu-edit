// SPEC-120 "Criteria" — functions/criteria.ts, called directly: parseCriterion
// over the whole criteria table (a number matches number cells equal at 15
// digits and numeric text; a boolean that boolean; an error that code; "" /
// "=" / "<>" select empty cells, "" results and non-empty cells; any other text
// is a wildcard pattern, case-insensitive unless the .ods says otherwise;
// < <= > >= order number cells numerically and text cells through the
// collator; <> is the negation of = over every cell, empty ones included) and
// its matchesEmpty; wildcardMatch and hasWildcards (`*`, `?` one UTF-16 unit,
// the `~*` `~?` `~~` escapes); textMatcher under the .ods settings ("plain"
// literal, "regex" literal or uncomputable, partial-cell matching);
// searchText; lookupMatcher (empty cells never match); and the pathological
// pattern `*a*a*a*a*a*b` against 10,000 a's finishing in linear-ish time.
// Expected values follow the spec's table; "Excel" marks Microsoft's examples.
import { describe, expect, it } from "vitest";
import { SPREADSHEET_LOCALES } from "../../numberFormat/locale";
import type { EvalSettings } from "../evaluate";
import { UncomputableError } from "../registry";
import { EMPTY, type ScalarValue } from "../values";
import { hasWildcards, lookupMatcher, parseCriterion, searchText, textMatcher, wildcardMatch } from "./criteria";

/** An xlsx workbook's settings, en-US. */
const XLSX: EvalSettings = {
  locale: SPREADSHEET_LOCALES["en-US"],
  dateSystem: "1900",
  caseSensitive: false,
  criteria: "wildcards",
  wholeCell: true,
  dateQuirks: "excel",
};
const GERMAN: EvalSettings = { ...XLSX, locale: SPREADSHEET_LOCALES["de-DE"] };
const CASE_SENSITIVE: EvalSettings = { ...XLSX, caseSensitive: true };
const PLAIN = { criteria: "plain", caseSensitive: false } as const;
const REGEX = { criteria: "regex", caseSensitive: false } as const;

const NA: ScalarValue = { error: "#N/A" };
const DIV0: ScalarValue = { error: "#DIV/0!" };

/** One cell of every kind the criteria table tells apart. */
const CELLS: ScalarValue[] = [EMPTY, "", 0, 1, 5, 6, "5", true, false, "apple", "Apple", "banana", NA, DIV0];
const except = (...left: ScalarValue[]): ScalarValue[] => CELLS.filter((cell) => !left.includes(cell));

/** The cells a criterion matches, in order. */
function matching(criterion: ScalarValue, cells: ScalarValue[] = CELLS, settings: EvalSettings = XLSX): ScalarValue[] {
  const test = parseCriterion(criterion, settings);
  return cells.filter((cell) => test.matches(cell));
}

/** What `run` throws (undefined when it returns). */
function thrown(run: () => unknown): unknown {
  try {
    run();
  } catch (error) {
    return error;
  }
  return undefined;
}

describe("parseCriterion", () => {
  it.each<[ScalarValue, ScalarValue[]]>([
    // A number, or text valueFromText reads as one: number cells and numeric text.
    [5, [5, "5"]],
    ["5", [5, "5"]],
    ["=5", [5, "5"]],
    ["<>5", except(5, "5")],
    [">5", [6]], // ordering reads number cells only
    [">=5", [5, 6]],
    ["<5", [0, 1]],
    ["<=5", [0, 1, 5]],
    [0, [0]], // neither an empty cell nor ""
    [EMPTY, [0]], // an empty-cell criterion is the number 0
    // Booleans.
    [true, [true]],
    ["TRUE", [true]],
    ["FALSE", [false]],
    ["<>TRUE", except(true)],
    [">FALSE", [true]],
    ["<TRUE", [false]],
    // Error literals.
    [NA, [NA]],
    ["#N/A", [NA]],
    ["#DIV/0!", [DIV0]],
    ["<>#N/A", except(NA)],
    [">#N/A", []],
    // Empties.
    ["", [EMPTY, ""]],
    ["=", [EMPTY]],
    ["<>", except(EMPTY)],
    // Other text: text cells, wildcards, case-insensitively.
    ["apple", ["apple", "Apple"]],
    ["=APPLE", ["apple", "Apple"]],
    ["a*", ["apple", "Apple"]],
    ["?pple", ["apple", "Apple"]],
    ["*an*", ["banana"]],
    ["*", ["", "5", "apple", "Apple", "banana"]],
    ["<>apple", except("apple", "Apple")],
    ["<>*", except("", "5", "apple", "Apple", "banana")],
  ])("%j matches %j", (criterion, expected) => {
    expect(matching(criterion)).toEqual(expected);
  });

  it.each<[string, ScalarValue[]]>([
    ["<b", ["apple", "Apple"]],
    [">=b", ["banana", "Banana", "cherry"]],
    ["<=banana", ["apple", "Apple", "banana", "Banana"]],
    [">banana", ["cherry"]],
  ])("%s orders text cells through the collator, case-insensitively: %j", (criterion, expected) => {
    expect(matching(criterion, [EMPTY, 1, true, "apple", "Apple", "banana", "Banana", "cherry"])).toEqual(expected);
  });

  it("compares numbers at 15 significant digits", () => {
    expect(parseCriterion(0.3, XLSX).matches(0.1 + 0.2)).toBe(true);
    expect(parseCriterion("=0.3", XLSX).matches(0.1 + 0.2)).toBe(true);
    expect(parseCriterion(0.3, XLSX).matches(0.31)).toBe(false);
  });

  it("reads dates, grouped numbers and decimals through valueFromText in the user's locale", () => {
    expect(matching("10/8/2026", [46_303, "10/8/2026", "2026-10-08", 46_304])).toEqual([46_303, "10/8/2026", "2026-10-08"]);
    expect(matching(">=10/8/2026", [46_302, 46_303, 46_304, "10/8/2026"])).toEqual([46_303, 46_304]);
    expect(matching("1,000", [1_000, "1000", "1,000", 1])).toEqual([1_000, "1000", "1,000"]);
    expect(matching("1,5", [1.5, "1,5"], GERMAN)).toEqual([1.5, "1,5"]);
    expect(matching("1,5", [1.5, "1,5"])).toEqual(["1,5"]); // en-US: not a number, so plain text
    expect(matching("50%", [0.5, "50%", 50])).toEqual([0.5, "50%"]);
    expect(matching(">=50%", [0.4, 0.5, 0.6])).toEqual([0.5, 0.6]);
    expect(matching("$4", [4, "4", "$4", 5])).toEqual([4, "4", "$4"]);
    expect(matching("12:00", [0.5, "12:00", 12])).toEqual([0.5, "12:00"]);
  });

  it("matches a wildcard pattern against text cells only, never numbers", () => {
    expect(matching("5*", [5, "5", 50, "50", "abc"])).toEqual(["5", "50"]);
  });

  it("matches text case-sensitively when an .ods says so", () => {
    expect(matching("apple", CELLS, CASE_SENSITIVE)).toEqual(["apple"]);
    expect(matching("A*", CELLS, CASE_SENSITIVE)).toEqual(["Apple"]);
    expect(matching("<>apple", CELLS, CASE_SENSITIVE)).toEqual(except("apple"));
  });

  it("matches = text criteria as substrings when the whole cell need not match", () => {
    expect(matching("pp", ["apple", "banana", "pepper"], { ...XLSX, wholeCell: false })).toEqual(["apple", "pepper"]);
  });

  it.each<[ScalarValue, boolean]>([
    [EMPTY, false],
    [0, false],
    [5, false],
    ["", true],
    ["=", true],
    ["<>", false],
    ["<>5", true],
    ["<>TRUE", true],
    ["<>#N/A", true],
    ["<>apple", true],
    ["<>*", true],
    ["apple", false],
    ["*", false],
    [">5", false],
    ["<5", false],
    ["<b", false],
  ])("matchesEmpty for %j is %s", (criterion, expected) => {
    expect(parseCriterion(criterion, XLSX).matchesEmpty).toBe(expected);
  });
});

describe("wildcardMatch and hasWildcards", () => {
  it.each<[string, string, boolean]>([
    ["apple", "a*", true],
    ["apple", "*e", true],
    ["apple", "*", true],
    ["", "*", true],
    ["", "?", false],
    ["apple", "a?ple", true],
    ["apple", "a?le", false],
    ["apple", "?????", true],
    ["apple", "??????", false],
    ["apple", "a*p*e", true],
    ["apple", "a*x*e", false],
    ["banana", "*an*an*", true],
    ["apple", "apple*", true],
    ["apple", "apples", false],
    ["APPLE", "a*e", true], // case-insensitive by default
    ["École", "é*", true],
    ["ecole", "é*", false], // accents still differ
    ["a*c", "a~*c", true],
    ["abc", "a~*c", false],
    ["a?c", "a~?c", true],
    ["abc", "a~?c", false],
    ["a~c", "a~~c", true],
    ["a~~c", "a~~c", false],
    ["😀", "?", false], // ? is one UTF-16 unit, and 😀 is two
    ["😀", "??", true],
  ])("%j against %j is %s", (text, pattern, expected) => {
    expect(wildcardMatch(text, pattern)).toBe(expected);
  });

  it("matches case-sensitively when asked", () => {
    expect(wildcardMatch("APPLE", "a*e", true)).toBe(false);
    expect(wildcardMatch("apple", "a*e", true)).toBe(true);
  });

  it.each<[string, boolean]>([
    ["a*", true],
    ["a?", true],
    ["*", true],
    ["~~*", true], // an escaped tilde, then a wildcard
    ["abc", false],
    ["a~*", false],
    ["a~?", false],
    ["~*~?", false],
    ["", false],
  ])("hasWildcards(%j) is %s", (pattern, expected) => {
    expect(hasWildcards(pattern)).toBe(expected);
  });
});

describe("textMatcher", () => {
  it("matches the whole text under wildcards, escapes resolved", () => {
    expect(textMatcher("a?c", XLSX)("ABC")).toBe(true);
    expect(textMatcher("a?c", XLSX)("abcd")).toBe(false);
    expect(textMatcher("a~*", XLSX)("a*")).toBe(true);
    expect(textMatcher("a~*", XLSX)("ab")).toBe(false);
    expect(textMatcher("pp", XLSX)("apple")).toBe(false);
  });

  it("reads a plain criterion as literal text: no wildcards, no escapes", () => {
    expect(textMatcher("a*", PLAIN)("a*")).toBe(true);
    expect(textMatcher("a*", PLAIN)("A*")).toBe(true);
    expect(textMatcher("a*", PLAIN)("apple")).toBe(false);
    expect(textMatcher("a~*", PLAIN)("a~*")).toBe(true);
    expect(textMatcher("a~*", PLAIN)("a*")).toBe(false);
  });

  it("reads a regex criterion with none of .*+?()[]{}|^$\\ literally", () => {
    expect(textMatcher("apple", REGEX)("Apple")).toBe(true);
    expect(textMatcher("apple", REGEX)("apples")).toBe(false);
  });

  it.each([".", "*", "+", "?", "(", ")", "[", "]", "{", "}", "|", "^", "$", "\\"])(
    "throws UncomputableError (regex) for a regex criterion holding %s",
    (special) => {
      const error = thrown(() => textMatcher(`a${special}b`, REGEX));
      expect(error).toBeInstanceOf(UncomputableError);
      expect((error as UncomputableError).reason).toEqual({ kind: "regex" });
    },
  );

  it("makes a regex criterion with metacharacters uncomputable through parseCriterion too", () => {
    const error = thrown(() => parseCriterion("app.*", { ...XLSX, criteria: "regex" }));
    expect((error as UncomputableError).reason).toEqual({ kind: "regex" });
  });

  it("matches substrings when partial", () => {
    expect(textMatcher("pp", XLSX, true)("apple")).toBe(true);
    expect(textMatcher("pp", XLSX, true)("banana")).toBe(false);
    expect(textMatcher("p?l", XLSX, true)("APPLE")).toBe(true);
    expect(textMatcher("PP", PLAIN, true)("apple")).toBe(true);
    expect(textMatcher("pp", REGEX, true)("apple")).toBe(true);
  });

  it("keeps a lone trailing ~ literal when matching partially, as whole-cell matching does", () => {
    // Only ~* ~? ~~ are escapes, so "a*~" is a, any run, then ~.
    expect(textMatcher("a*~", XLSX)("ab~")).toBe(true);
    expect(textMatcher("a*~", XLSX, true)("xab~")).toBe(true);
  });

  it("honours caseSensitive", () => {
    const sensitive = { criteria: "wildcards", caseSensitive: true } as const;
    expect(textMatcher("Apple", sensitive)("apple")).toBe(false);
    expect(textMatcher("Apple", sensitive)("Apple")).toBe(true);
    expect(textMatcher("A*", sensitive)("apple")).toBe(false);
  });
});

describe("searchText", () => {
  it.each<[string, string, number, number]>([
    ["Statements", "e", 5, 6], // Excel: SEARCH("e","Statements",6) is 7
    ["Profit Margin", "margin", 0, 7], // Excel: SEARCH("margin","Profit Margin") is 8
    ["Profit Margin", " ", 0, 6], // Excel: MID(A3,SEARCH(" ",A3)+1,4) is "Marg"
    ['The "boss" is here.', '"', 0, 4], // Excel: SEARCH("""",A5) is 5
    ["Profit Margin", "m?rg*", 0, 7],
    ["bank", "?a", 0, 0],
    ["2*3=6", "~*", 0, 1],
    ["apple", "x", 0, -1],
    ["apple", "p", 3, -1],
  ])("finds %j in %j from %d at %d (0-based)", (within, find, start, expected) => {
    expect(searchText(within, find, start, XLSX)).toBe(expected);
  });

  it("keeps a lone trailing ~ literal: ?~ is one unit, then ~", () => {
    expect(searchText("xa~y", "?~", 0, XLSX)).toBe(1);
  });

  it("treats * literally in plain mode, as any run under wildcards", () => {
    expect(searchText("2*3", "*", 0, PLAIN)).toBe(1);
    expect(searchText("2*3", "*", 0, XLSX)).toBe(0);
  });
});

describe("lookupMatcher", () => {
  it("matches a number to numbers equal at 15 digits — not numeric text, logicals or empty cells", () => {
    const match = lookupMatcher(0.3, XLSX);
    expect([0.3, 0.1 + 0.2, 0.31, "0.3", false, EMPTY].map((cell) => match(cell))).toEqual([true, true, false, false, false, false]);
    expect(lookupMatcher("2", XLSX)(2)).toBe(false);
  });

  it("matches a logical to that logical only", () => {
    const match = lookupMatcher(true, XLSX);
    expect([true, false, 1, "TRUE", EMPTY].map((cell) => match(cell))).toEqual([true, false, false, false, false]);
  });

  it("matches text whole-cell with wildcards, case-insensitively unless caseSensitive", () => {
    const match = lookupMatcher("B*", XLSX); // Excel: MATCH("B*",{"apple","banana"},0) is 2
    expect(["apple", "banana", "Banana", "", EMPTY].map((cell) => match(cell))).toEqual([false, true, true, false, false]);
    expect(lookupMatcher("Apple", CASE_SENSITIVE)("apple")).toBe(false);
    expect(lookupMatcher("Apple", CASE_SENSITIVE)("Apple")).toBe(true);
  });

  it.each<ScalarValue>(["*", "", 0, false, EMPTY])("never matches an empty cell, looking up %j", (lookup) => {
    expect(lookupMatcher(lookup, XLSX)(EMPTY)).toBe(false);
  });
});

describe("the pathological pattern", () => {
  it("rejects *a*a*a*a*a*b against 10,000 a's in well under a second", () => {
    const text = "a".repeat(10_000);
    const pattern = "*a*a*a*a*a*b";
    const start = performance.now();
    expect(wildcardMatch(text, pattern)).toBe(false);
    expect(textMatcher(pattern, XLSX)(text)).toBe(false);
    expect(textMatcher(pattern, XLSX, true)(text)).toBe(false);
    expect(parseCriterion(pattern, XLSX).matches(text)).toBe(false);
    expect(performance.now() - start).toBeLessThan(500);
    expect(wildcardMatch(`${text}b`, pattern)).toBe(true);
  });
});
