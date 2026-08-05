import { describe, it, expect } from "vitest";
import type { DocumentModel, Paragraph, Run } from "./documentModel";
import { blockText } from "./textEditing";
import {
  compileQuery,
  escapeRegExp,
  findMatches,
  INVALID_REGEX_MESSAGE,
  indexAtOrAfter,
  isQueryError,
  matchSelection,
  replaceMatch,
  replaceMatches,
  type CompiledQuery,
  type FindOptions,
} from "./findReplace";

const DEFAULT_OPTIONS: FindOptions = { matchCase: false, wholeWords: false, regex: false };

function options(overrides: Partial<FindOptions> = {}): FindOptions {
  return { ...DEFAULT_OPTIONS, ...overrides };
}

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function docOf(...blocks: Paragraph[]): DocumentModel {
  return {
    id: "doc",
    title: "t",
    sections: [{ blocks }],
    defaultStyles: {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 },
      fontFamily: "Calibri",
      fontSizePt: 11,
    },
  };
}

function compiled(query: string, opts: FindOptions = DEFAULT_OPTIONS): CompiledQuery {
  const result = compileQuery(query, opts);
  if (result === null || isQueryError(result)) {
    throw new Error("expected a compiled query");
  }
  return result;
}

function matchTexts(model: DocumentModel, query: string, opts: FindOptions = DEFAULT_OPTIONS): string[] {
  return findMatches(model, compiled(query, opts)).map((match) => {
    const text = blockText(model.sections[match.section].blocks[match.block]);
    return text.slice(match.start, match.end);
  });
}

describe("compileQuery", () => {
  it("returns null for an empty query", () => {
    expect(compileQuery("", DEFAULT_OPTIONS)).toBeNull();
  });

  it("compiles a literal query verbatim, ignoring regex metacharacters", () => {
    const model = docOf(paragraph(["a.b axb a.b"]));
    expect(matchTexts(model, "a.b")).toEqual(["a.b", "a.b"]);
  });

  it("is case-insensitive by default and case-sensitive with matchCase", () => {
    const model = docOf(paragraph(["Cat cat CAT"]));
    expect(matchTexts(model, "cat")).toEqual(["Cat", "cat", "CAT"]);
    expect(matchTexts(model, "cat", options({ matchCase: true }))).toEqual(["cat"]);
  });

  it("matches whole words only when requested", () => {
    const model = docOf(paragraph(["cat category scatter cat"]));
    expect(matchTexts(model, "cat", options({ wholeWords: true }))).toEqual(["cat", "cat"]);
  });

  it("interprets the query as a regular expression in regex mode", () => {
    const model = docOf(paragraph(["a1 b2 c3"]));
    expect(matchTexts(model, "[a-z]\\d", options({ regex: true }))).toEqual(["a1", "b2", "c3"]);
  });

  it("reports an inline error for an invalid regular expression", () => {
    const result = compileQuery("(unclosed", options({ regex: true }));
    expect(result).not.toBeNull();
    expect(isQueryError(result as never)).toBe(true);
    expect((result as { error: string }).error).toBe(INVALID_REGEX_MESSAGE);
  });

  it("combines whole words with regex mode", () => {
    const model = docOf(paragraph(["cat cats cat"]));
    expect(matchTexts(model, "ca.", options({ regex: true, wholeWords: true }))).toEqual(["cat", "cat"]);
  });
});

describe("escapeRegExp", () => {
  it("escapes every regex metacharacter", () => {
    expect(escapeRegExp("a.b*c")).toBe("a\\.b\\*c");
    expect(new RegExp(escapeRegExp("(x)")).test("(x)")).toBe(true);
  });
});

describe("findMatches", () => {
  it("finds matches across multiple paragraphs in document order", () => {
    const model = docOf(paragraph(["one two"]), paragraph(["two three"]));
    const found = findMatches(model, compiled("two"));
    expect(found).toEqual([
      { section: 0, block: 0, start: 4, end: 7 },
      { section: 0, block: 1, start: 0, end: 3 },
    ]);
  });

  it("does not span a paragraph boundary", () => {
    const model = docOf(paragraph(["foo"]), paragraph(["bar"]));
    expect(matchTexts(model, "foobar")).toEqual([]);
  });

  it("skips zero-length matches instead of looping forever", () => {
    const model = docOf(paragraph(["abc"]));
    // `a*` matches empty strings at every position; none should be recorded.
    expect(findMatches(model, compiled("a*", options({ regex: true })))).toEqual([
      { section: 0, block: 0, start: 0, end: 1 },
    ]);
  });

  it("finds matches spanning run boundaries", () => {
    const model = docOf(paragraph(["cat", { bold: true }], ["egory"]));
    expect(matchTexts(model, "category")).toEqual(["category"]);
  });
});

describe("matchSelection", () => {
  it("maps a match to its two caret ends", () => {
    expect(matchSelection({ section: 1, block: 2, start: 3, end: 8 })).toEqual({
      anchor: { section: 1, block: 2, offset: 3 },
      focus: { section: 1, block: 2, offset: 8 },
    });
  });
});

describe("replaceMatches", () => {
  it("replaces every match and reports through the caller's count", () => {
    const model = docOf(paragraph(["cat cat cat"]));
    const matches = findMatches(model, compiled("cat"));
    const result = replaceMatches(model, matches, "dog");
    expect(blockText(result.model.sections[0].blocks[0])).toBe("dog dog dog");
  });

  it("does not mutate the input model", () => {
    const model = docOf(paragraph(["cat cat"]));
    const matches = findMatches(model, compiled("cat"));
    replaceMatches(model, matches, "dog");
    expect(blockText(model.sections[0].blocks[0])).toBe("cat cat");
  });

  it("takes the formatting of the first run of the replaced text", () => {
    // "cat" is bold; the run to its left ("plain ") is not. The replacement
    // must be bold — the first replaced run's formatting, not the neighbour's.
    const model = docOf(paragraph(["plain "], ["cat", { bold: true }], [" plain"]));
    const matches = findMatches(model, compiled("cat"));
    const result = replaceMatches(model, matches, "dog");
    const runs = result.model.sections[0].blocks[0].runs;
    const dogRun = runs.find((run: Run) => run.text.includes("dog"));
    expect(dogRun?.formatting).toEqual({ bold: true });
  });

  it("uses the first replaced run's formatting when a match spans two runs", () => {
    const model = docOf(paragraph(["ca", { italic: true }], ["t", { bold: true }]));
    const matches = findMatches(model, compiled("cat"));
    const result = replaceMatches(model, matches, "dog");
    const runs = result.model.sections[0].blocks[0].runs;
    expect(runs).toEqual([{ text: "dog", formatting: { italic: true } }]);
  });

  it("applies several matches in one paragraph without corrupting offsets", () => {
    const model = docOf(paragraph(["ab ab ab"]));
    const matches = findMatches(model, compiled("ab"));
    const result = replaceMatches(model, matches, "xyz");
    expect(blockText(result.model.sections[0].blocks[0])).toBe("xyz xyz xyz");
  });

  it("supports deletion with an empty replacement", () => {
    const model = docOf(paragraph(["a-b-c"]));
    const matches = findMatches(model, compiled("-"));
    const result = replaceMatches(model, matches, "");
    expect(blockText(result.model.sections[0].blocks[0])).toBe("abc");
  });

  it("returns a caret just after the first replacement", () => {
    const model = docOf(paragraph(["cat cat"]));
    const matches = findMatches(model, compiled("cat"));
    const result = replaceMatches(model, matches, "hippo");
    expect(result.selection.focus).toEqual({ section: 0, block: 0, offset: 5 });
  });
});

describe("replaceMatch", () => {
  it("replaces only the given match", () => {
    const model = docOf(paragraph(["cat cat cat"]));
    const matches = findMatches(model, compiled("cat"));
    const result = replaceMatch(model, matches[1], "dog");
    expect(blockText(result.model.sections[0].blocks[0])).toBe("cat dog cat");
  });
});

describe("indexAtOrAfter", () => {
  const matches = [
    { section: 0, block: 0, start: 0, end: 3 },
    { section: 0, block: 1, start: 5, end: 8 },
    { section: 0, block: 2, start: 2, end: 5 },
  ];

  it("returns -1 when there are no matches", () => {
    expect(indexAtOrAfter([], { section: 0, block: 0, offset: 0 })).toBe(-1);
  });

  it("finds the first match at or after the anchor", () => {
    expect(indexAtOrAfter(matches, { section: 0, block: 1, offset: 0 })).toBe(1);
    expect(indexAtOrAfter(matches, { section: 0, block: 1, offset: 5 })).toBe(1);
  });

  it("wraps to the first match when the anchor is past the last match", () => {
    expect(indexAtOrAfter(matches, { section: 0, block: 9, offset: 0 })).toBe(0);
  });
});
