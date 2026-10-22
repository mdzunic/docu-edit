// SPEC-127 "The assists" and "Function metadata": functionSignature over
// fixed, optional, single- and double-group repeats; argumentAt past the
// repeating group and past maxArgs; callAtCaret over nesting, plain
// parentheses, unions, array constants, quoted sheet names, strings holding
// "," and "(", unterminated strings and empty text; completion contexts; and
// the dialog's search ranking.
import { describe, expect, it } from "vitest";
import { argumentAt, callAtCaret, completionAt, functionSignature, functionsStartingWith, searchFunctions, signatureParts } from "./functionHelp";
import { getFunction, type FunctionSpec } from "./registry";

function spec(name: string): FunctionSpec {
  const found = getFunction(name);
  if (!found) {
    throw new Error(`${name} is not registered`);
  }
  return found;
}

/** The caret at the end of the text. */
const atEnd = (text: string) => callAtCaret(text, text.length);

describe("functionSignature", () => {
  it.each([
    ["PI", "PI()"],
    ["ROUND", "ROUND(number, num_digits)"],
    ["PMT", "PMT(rate, nper, pv, [fv], [type])"],
    ["SUM", "SUM(number1, [number2], …)"],
    ["SUMIFS", "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)"],
    ["XLOOKUP", "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])"],
  ])("%s → %s", (name, signature) => {
    expect(functionSignature(spec(name))).toBe(signature);
  });

  it("repeats the group up to the current argument's instance, marking it", () => {
    const parts = signatureParts(spec("SUMIFS"), 5);
    expect(parts.map((part) => part.text).join("")).toBe("SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], [criteria_range3, criteria3], …)");
    expect(parts.filter((part) => part.current).map((part) => part.text)).toEqual(["criteria_range3"]);
    expect(signatureParts(spec("ROUND"), 2).some((part) => part.current)).toBe(false);
  });
});

describe("argumentAt", () => {
  it("names fixed and optional arguments, and null past maxArgs", () => {
    expect([0, 1, 2, 3, 4, 5].map((i) => argumentAt(spec("PMT"), i))).toEqual(["rate", "nper", "pv", "fv", "type", null]);
  });

  it("numbers a single-argument group past its end (number7) and stops at 255", () => {
    expect(argumentAt(spec("SUM"), 0)).toBe("number1");
    expect(argumentAt(spec("SUM"), 6)).toBe("number7");
    expect(argumentAt(spec("SUM"), 254)).toBe("number255");
    expect(argumentAt(spec("SUM"), 255)).toBeNull();
  });

  it("cycles a two-argument group (criteria_range3, criteria3)", () => {
    expect([3, 4, 5, 6].map((i) => argumentAt(spec("SUMIFS"), i))).toEqual(["criteria_range2", "criteria2", "criteria_range3", "criteria3"]);
  });
});

describe("callAtCaret", () => {
  it.each<[string, { name: string; argIndex: number } | null]>([
    ["=SUM(", { name: "SUM", argIndex: 0 }],
    ["=SUM(1,", { name: "SUM", argIndex: 1 }],
    ["=SUM(1,ROUND(2,", { name: "ROUND", argIndex: 1 }],
    ["=SUM(1,ROUND(2,3),", { name: "SUM", argIndex: 2 }],
    ["=SUM((1+2),", { name: "SUM", argIndex: 1 }],
    ["=SUM((A1,B1),", { name: "SUM", argIndex: 1 }],
    ["=SUM({1,2;3,4},", { name: "SUM", argIndex: 1 }],
    ["=SUM({1,2", { name: "SUM", argIndex: 0 }],
    ["=SUM('a,b'!A1,", { name: "SUM", argIndex: 1 }],
    ['=IF(",(",', { name: "IF", argIndex: 1 }],
    ['=IF("a,', null],
    ["=NOSUCH(1,", null],
    ["=SUM(1,NOSUCH(2,", { name: "SUM", argIndex: 1 }],
    ["", null],
    ["=", null],
  ])("%s → %j", (text, expected) => {
    expect(atEnd(text)).toEqual(expected);
  });

  it("reads only the text before the caret", () => {
    expect(callAtCaret("=ROUND(1,2)", 7)).toEqual({ name: "ROUND", argIndex: 0 });
    expect(callAtCaret("=ROUND(1,2)", 11)).toBeNull();
  });
});

describe("completion contexts", () => {
  it.each<[string, { prefix: string; start: number } | null]>([
    ["=XL", { prefix: "XL", start: 1 }],
    ["=1+ro", { prefix: "ro", start: 3 }],
    ["=SUM(A1,av", { prefix: "av", start: 8 }],
    ["=SUM(ro", { prefix: "ro", start: 5 }],
    ["=LOG1", { prefix: "LOG1", start: 1 }],
    ['="XL', null],
    ['=IF("a",XL', { prefix: "XL", start: 8 }],
    ["=SUM(A1:B1)", null],
    ["=XL ", null],
    ["XL", null],
    ["=Sheet1!A", null],
    ["=$A", null],
  ])("%s → %j", (text, expected) => {
    expect(completionAt(text, text.length)).toEqual(expected);
  });

  it("lists registered names starting with the prefix, case-insensitively, alphabetically", () => {
    expect(functionsStartingWith("xl")).toEqual(["XLOOKUP"]);
    expect(functionsStartingWith("ROU")).toEqual(["ROUND", "ROUNDDOWN", "ROUNDUP"]);
    expect(functionsStartingWith("SEQ")).toEqual([]);
  });
});

describe("searchFunctions", () => {
  it("lists a category alphabetically for the empty query, every registered name for All", () => {
    const financial = searchFunctions("", "Financial");
    expect(financial).toEqual(financial.toSorted());
    expect(financial.every((name) => getFunction(name)?.category === "Financial")).toBe(true);
    expect(searchFunctions("", null).length).toBeGreaterThan(300);
  });

  it("ranks name prefix, then name substring, then summary words, alphabetically within each", () => {
    const found = searchFunctions("count", null);
    const prefix = found.filter((name) => name.startsWith("COUNT"));
    expect(found.slice(0, prefix.length)).toEqual(prefix.toSorted());
    expect(searchFunctions("payment", "Financial")).toContain("PMT");
    expect(searchFunctions("no such words here", null)).toEqual([]);
  });

  it("never offers a dynamic-array function", () => {
    expect(searchFunctions("SEQUENCE", null)).not.toContain("SEQUENCE");
    expect(searchFunctions("FILTER", null)).not.toContain("FILTER");
  });
});
