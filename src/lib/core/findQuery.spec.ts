// The product-neutral find query (SPEC-098's move into core, performed by
// SPEC-134): compileQuery under each option, the empty and invalid cases,
// escapeRegExp and isQueryError — the same behaviour findReplace.spec.ts
// pins through the text editor's re-export.
import { describe, expect, it } from "vitest";
import { compileQuery, escapeRegExp, INVALID_REGEX_MESSAGE, isQueryError, type CompiledQuery } from "./findQuery";

const OFF = { matchCase: false, wholeWords: false, regex: false };

function compiled(query: string, options = OFF): RegExp {
  const result = compileQuery(query, options);
  if (result === null || isQueryError(result)) {
    throw new Error(`"${query}" did not compile`);
  }
  return (result as CompiledQuery).regExp;
}

describe("compileQuery", () => {
  it("returns null for an empty query: there is nothing to search for", () => {
    expect(compileQuery("", OFF)).toBeNull();
    expect(compileQuery("", { matchCase: true, wholeWords: true, regex: true })).toBeNull();
  });

  it("is global and case-insensitive by default, case-sensitive with Match case", () => {
    expect(compiled("cat").flags).toBe("gi");
    expect(compiled("cat", { ...OFF, matchCase: true }).flags).toBe("g");
    expect("Cat CAT cat".match(compiled("cat"))).toEqual(["Cat", "CAT", "cat"]);
    expect("Cat CAT cat".match(compiled("cat", { ...OFF, matchCase: true }))).toEqual(["cat"]);
  });

  it("treats a non-regex query literally — every metacharacter matches itself", () => {
    expect("a.b axb".match(compiled("a.b"))).toEqual(["a.b"]);
    expect("1+1=2".match(compiled("1+1"))).toEqual(["1+1"]);
    expect("*?".match(compiled("*?"))).toEqual(["*?"]);
    expect("$1 (x)".match(compiled("(x)"))).toEqual(["(x)"]);
  });

  it("wraps Whole words in word boundaries, in every mode", () => {
    expect(compiled("cat", { ...OFF, wholeWords: true }).source).toBe("\\b(?:cat)\\b");
    expect("cat concat cats cat.".match(compiled("cat", { ...OFF, wholeWords: true }))).toEqual(["cat", "cat"]);
    expect("ab abc".match(compiled("a.", { ...OFF, regex: true, wholeWords: true }))).toEqual(["ab"]);
  });

  it("compiles a regular expression as written in regex mode", () => {
    expect("a1 b22 c".match(compiled("[a-z]\\d+", { ...OFF, regex: true }))).toEqual(["a1", "b22"]);
  });

  it("reports an invalid regular expression instead of throwing", () => {
    const result = compileQuery("(unclosed", { ...OFF, regex: true });
    expect(result).toEqual({ error: INVALID_REGEX_MESSAGE });
    expect(result !== null && isQueryError(result)).toBe(true);
    // The same text is literal — and valid — outside regex mode.
    expect("(unclosed".match(compiled("(unclosed"))).toEqual(["(unclosed"]);
  });
});

describe("escapeRegExp and isQueryError", () => {
  it("escapes every regex metacharacter", () => {
    const literal = ".*+?^${}()|[]\\";
    expect(new RegExp(escapeRegExp(literal)).test(literal)).toBe(true);
    expect(escapeRegExp("plain text")).toBe("plain text");
  });

  it("narrows only the error case", () => {
    expect(isQueryError({ error: INVALID_REGEX_MESSAGE })).toBe(true);
    expect(isQueryError({ regExp: /x/g })).toBe(false);
    expect(INVALID_REGEX_MESSAGE).toBe("Invalid regular expression");
  });
});
