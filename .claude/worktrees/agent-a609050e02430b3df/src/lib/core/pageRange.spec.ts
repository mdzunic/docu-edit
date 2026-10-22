import { describe, expect, it } from "vitest";
import { parseNumberRange, rangeHint } from "./pageRange";

describe("parseNumberRange", () => {
  it("parses single numbers, ranges and comma lists, deduplicated and sorted", () => {
    expect(parseNumberRange("3", 10).pages).toEqual([3]);
    expect(parseNumberRange("2-4", 10).pages).toEqual([2, 3, 4]);
    expect(parseNumberRange("5, 1-2, 2", 10).pages).toEqual([1, 2, 5]);
    expect(parseNumberRange(" 1 - 3 ", 10).pages).toEqual([1, 2, 3]);
  });

  it("defaults to the page noun, wording every error exactly as SPEC-034 did", () => {
    expect(parseNumberRange("", 10).error).toBe('Enter a page range. Use page numbers like "1-3, 5"');
    expect(parseNumberRange("abc", 10).error).toBe('"abc" is not a page number. Use page numbers like "1-3, 5"');
    expect(parseNumberRange("4-2", 10).error).toBe('"4-2" starts after it ends');
    expect(parseNumberRange("11", 10).error).toBe("Pages go from 1 to 10");
    expect(parseNumberRange("2", 1).error).toBe("This document only has 1 page");
  });

  it("words every error with the noun it is given (the deck dialog passes \"slide\")", () => {
    expect(parseNumberRange("", 4, "slide").error).toBe('Enter a slide range. Use slide numbers like "1-3, 5"');
    expect(parseNumberRange("x", 4, "slide").error).toBe('"x" is not a slide number. Use slide numbers like "1-3, 5"');
    expect(parseNumberRange("0", 4, "slide").error).toBe("Slides go from 1 to 4");
    expect(parseNumberRange("2", 1, "slide").error).toBe("This document only has 1 slide");
    expect(parseNumberRange("1, 3-4", 4, "slide").pages).toEqual([1, 3, 4]);
  });

  it("rangeHint names the noun", () => {
    expect(rangeHint()).toBe('Use page numbers like "1-3, 5"');
    expect(rangeHint("slide")).toBe('Use slide numbers like "1-3, 5"');
  });
});
