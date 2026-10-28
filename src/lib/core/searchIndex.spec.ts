import { describe, expect, it } from "vitest";
import {
  BM25_B,
  BM25_K1,
  INDEX_VERSION,
  MAX_SNIPPETS_PER_DOC,
  MAX_TOKEN_LENGTH,
  TITLE_BOOST,
  bm25Idf,
  bm25Term,
  buildIndex,
  findTextFor,
  indexDocument,
  isSearchQueryError,
  mergeIndex,
  parseQuery,
  removeFromIndex,
  search,
  snippet,
  snippets,
  titleSegments,
  tokenise,
  type IndexedDoc,
  type Query,
  type QueryError,
  type SearchHit,
} from "./searchIndex";

const words = (text: string): string[] => tokenise(text).map((token) => token.token);

function query(input: string): Query {
  const parsed = parseQuery(input);
  if (isSearchQueryError(parsed)) {
    throw new Error(`unexpected QueryError: ${parsed.error} at ${parsed.position}`);
  }
  return parsed;
}

function queryError(input: string): QueryError {
  const parsed = parseQuery(input);
  if (!isSearchQueryError(parsed)) {
    throw new Error(`expected a QueryError for ${JSON.stringify(input)}`);
  }
  return parsed;
}

function doc(id: string, title: string, body: string, lastModified = 1, facets?: Record<string, string[]>): IndexedDoc {
  return { id, title, body, lastModified, ...(facets ? { facets } : {}) };
}

const ids = (hits: SearchHit[]): string[] => hits.map((hit) => hit.id);

describe("tokenise — each rule (a silent change invalidates every stored index)", () => {
  it("lowercases", () => {
    expect(words("Rotterdam HARBOUR Figures")).toEqual(["rotterdam", "harbour", "figures"]);
  });

  it("NFKD-normalises and strips combining marks, so café matches cafe", () => {
    const precomposed = "caf\u00e9";
    const decomposed = "cafe\u0301";
    expect(words(precomposed)).toEqual(["cafe"]);
    expect(words(decomposed)).toEqual(["cafe"]);
    expect(words("Ångström naïve Œuvre")).toEqual(["angstrom", "naive", "œuvre"]);
    // Compatibility forms decompose too: a ligature, full-width letters.
    expect(words("\ufb01ne \uff21\uff22\uff23")).toEqual(["fine", "abc"]);
  });

  it("splits on anything that is not a letter or digit", () => {
    expect(words("rotterdam,figures;q3/2026 (draft)!")).toEqual(["rotterdam", "figures", "q3", "2026", "draft"]);
    expect(words("tab\there\nnew line")).toEqual(["tab", "here", "new", "line"]);
    expect(words("Straße über ünïcödé")).toEqual(["straße", "uber", "unicode"]);
  });

  it("keeps ' and - inside a word, and only there (hyphenation)", () => {
    expect(words("don't e-mail well-known")).toEqual(["don't", "e-mail", "well-known"]);
    expect(words("'quoted' -dash- trailing- -leading")).toEqual(["quoted", "dash", "trailing", "leading"]);
    expect(words("a--b x'-y")).toEqual(["a", "b", "x", "y"]);
    // Typographic apostrophes and Unicode hyphens fold to ' and -.
    expect(words("don\u2019t co\u2010op")).toEqual(["don't", "co-op"]);
  });

  it("keeps numbers whole: . and , between two digits stay inside the token", () => {
    expect(words("3.14 1,250 1,234.56 v1.2.3")).toEqual(["3.14", "1,250", "1,234.56", "v1.2.3"]);
    expect(words("in 2026. Then, 7")).toEqual(["in", "2026", "then", "7"]);
    expect(words("a.b 3.x")).toEqual(["a", "b", "3", "x"]);
  });

  it("caps tokens at 64 characters, offsets still covering the whole word", () => {
    const long = "x".repeat(100);
    const [token] = tokenise(long);
    expect(token.token).toBe("x".repeat(MAX_TOKEN_LENGTH));
    expect(token.start).toBe(0);
    expect(token.end).toBe(100);
    // A surrogate pair straddling the cut is never split (Gothic ahsa, a
    // letter outside the BMP with no case and no decomposition).
    const astral = `${"a".repeat(63)}${String.fromCodePoint(0x10330)}b`;
    expect(tokenise(astral)[0].token).toBe("a".repeat(63));
    // Compatibility forms that decompose to capitals still come out lowercase.
    expect(words(String.fromCodePoint(0x1d400, 0x1d401))).toEqual(["ab"]);
  });

  it("indexes a run of CJK or Kana as overlapping bigrams", () => {
    expect(tokenise("東京都庁")).toEqual([
      { token: "東京", start: 0, end: 2 },
      { token: "京都", start: 1, end: 3 },
      { token: "都庁", start: 2, end: 4 },
    ]);
    expect(words("東")).toEqual(["東"]);
    expect(words("カタカナ ひらがな")).toEqual(["カタ", "タカ", "カナ", "ひら", "らが", "がな"].map((bigram) => bigram.normalize("NFKD").replace(/\p{M}/gu, "")));
    // A number stays whole beside a CJK run; Latin words split from it.
    expect(words("2026年東京Tokyo")).toEqual(["2026", "年東", "東京", "tokyo"]);
  });

  it("reports offsets into the original text, across expansions", () => {
    expect(tokenise("\ufb01ne caf\u00e9")).toEqual([
      { token: "fine", start: 0, end: 3 },
      { token: "cafe", start: 4, end: 8 },
    ]);
  });
});

describe("parseQuery — the grammar", () => {
  it("reads space-separated words as AND clauses, the last one a prefix while typing", () => {
    const parsed = query("rotterdam figures");
    expect(parsed.terms.map((term) => [term.tokens, term.prefix, term.exclude])).toEqual([
      [["rotterdam"], false, false],
      [["figures"], true, false],
    ]);
    expect(parsed.facets).toEqual([]);
  });

  it("ends the prefix with an explicit trailing space", () => {
    expect(query("rotter").terms[0].prefix).toBe(true);
    expect(query("rotter ").terms[0].prefix).toBe(false);
  });

  it("reads a quoted phrase as adjacent tokens, never a prefix", () => {
    const [term] = query('"Rotterdam figures"').terms;
    expect(term).toMatchObject({ tokens: ["rotterdam", "figures"], phrase: true, prefix: false, text: "Rotterdam figures" });
  });

  it("reads a leading - as an exclusion (never a prefix)", () => {
    const parsed = query("harbour -draft");
    expect(parsed.terms[1]).toMatchObject({ tokens: ["draft"], exclude: true, prefix: false });
    expect(query('harbour -"first draft"').terms[1]).toMatchObject({ phrase: true, exclude: true });
  });

  it("reads title:, kind: and format: filters", () => {
    const parsed = query("title:report kind:presentation format:PPTX -format:odp");
    expect(parsed.terms).toHaveLength(1);
    expect(parsed.terms[0]).toMatchObject({ tokens: ["report"], field: "title" });
    expect(parsed.facets.map((facet) => [facet.facet, facet.value, facet.exclude, facet.prefix])).toEqual([
      ["kind", "presentation", false, false],
      ["format", "pptx", false, false],
      ["format", "odp", true, false],
    ]);
    expect(query("Title:report").terms[0].field).toBe("title");
    expect(query("kind:pres").facets[0].prefix).toBe(true);
  });

  it("matches a word the tokeniser splits like a phrase", () => {
    expect(query("U.S.A ").terms[0]).toMatchObject({ tokens: ["u", "s", "a"], phrase: true });
    expect(query("東京都 ").terms[0]).toMatchObject({ tokens: ["東京", "京都"], phrase: true });
  });

  it("reads an empty or blank input as an empty query", () => {
    expect(query("")).toEqual({ terms: [], facets: [] });
    expect(query("   ")).toEqual({ terms: [], facets: [] });
  });

  // Every error, each with the offending position — never a throw, never a
  // query that matches everything.
  it.each([
    ['"rotterdam figures', "never closed", 0],
    ['harbour ""', "empty", 8],
    ['" "', "empty", 0],
    ["harbour - draft", "needs a word", 8],
    ["harbour -", "needs a word", 8],
    ["--draft", "Only one", 1],
    ["title: report", "needs a value", 6],
    ["kind:", "needs a value", 5],
    ["title:-draft", "before the filter", 6],
    ['harb"our', "only start a phrase", 4],
    ['"a b"c', "space after", 5],
    ["!!! harbour", "no letters or digits", 0],
    ['harbour "?!"', "no letters or digits", 9],
    ["-draft", "only narrows", 0],
    ["-draft -kind:pdf", "only narrows", 0],
  ])("%j is a QueryError (%s) at %i", (input, message, position) => {
    const error = queryError(input);
    expect(error.error).toContain(message);
    expect(error.position).toBe(position);
  });
});

describe("search — BM25 against hand-computed values", () => {
  // N = 3. Body lengths 3, 2, 4 (avg 3); title lengths 2, 1, 1 (avg 4/3).
  const index = buildIndex([
    doc("d1", "Harbour report", "rotterdam figures rotterdam", 30),
    doc("d2", "Rotterdam", "amsterdam figures", 20),
    doc("d3", "Notes", "nothing here at all", 10),
  ]);

  it("states the constants", () => {
    expect([BM25_K1, BM25_B, TITLE_BOOST]).toEqual([1.2, 0.75, 3]);
  });

  it("scores body and title with the same formula, the title × 3, summed across terms", () => {
    // "rotterdam": df_body = 1 (d1), df_title = 1 (d2).
    const idf = Math.log(1 + (3 - 1 + 0.5) / (1 + 0.5));
    expect(bm25Idf(3, 1)).toBeCloseTo(idf, 12);
    // d1: body tf 2, dl 3, avgdl 3 → idf × 2 × 2.2 / (2 + 1.2 × 1).
    const d1 = (idf * 2 * 2.2) / (2 + 1.2 * (1 - 0.75 + 0.75 * (3 / 3)));
    // d2: title tf 1, dl 1, avgdl 4/3 → 3 × idf × 2.2 / (1 + 1.2 × (0.25 + 0.75 × 0.75)).
    const d2 = (3 * (idf * 1 * 2.2)) / (1 + 1.2 * (1 - 0.75 + 0.75 * (1 / (4 / 3))));
    const hits = search(index, query("rotterdam "));
    expect(ids(hits)).toEqual(["d2", "d1"]);
    expect(hits[0].score).toBeCloseTo(d2, 10);
    expect(hits[1].score).toBeCloseTo(d1, 10);
    expect(hits[0].titleMatches).toBe(1);
    expect(hits[1].bodyMatches).toBe(2);

    // Two terms sum: "figures" (df_body 2) is in d1 and d2.
    const idfFigures = Math.log(1 + (3 - 2 + 0.5) / (2 + 0.5));
    const d1Figures = (idfFigures * 1 * 2.2) / (1 + 1.2 * (0.25 + 0.75 * 1));
    const both = search(index, query("rotterdam figures "));
    expect(ids(both)).toEqual(["d2", "d1"]);
    expect(both[1].score).toBeCloseTo(d1 + d1Figures, 10);
    expect(bm25Term(0, idf, 3, 3)).toBe(0);
  });

  it("breaks ties by lastModified, newest first, then by id", () => {
    const tied = buildIndex([
      doc("b", "Same", "alpha beta", 5),
      doc("a", "Same", "alpha beta", 5),
      doc("c", "Same", "alpha beta", 9),
    ]);
    const hits = search(tied, query("alpha "));
    expect(hits.map((hit) => hit.score)).toEqual([hits[0].score, hits[0].score, hits[0].score]);
    expect(ids(hits)).toEqual(["c", "a", "b"]);
  });

  it("ANDs clauses, excludes, filters by title, kind and format", () => {
    const corpus = buildIndex([
      doc("text", "Harbour report", "The Rotterdam figures, final draft.", 4, { kind: ["text"], format: ["docx"] }),
      doc("deck", "Quarterly review", "Rotterdam growth slides", 3, { kind: ["presentation"], format: ["pptx"] }),
      doc("pdf", "Scan", "Rotterdam harbour", 2, { kind: ["media"], format: ["pdf"] }),
    ]);
    // "harbour" is in the text's TITLE (× 3) but only the PDF's body.
    expect(ids(search(corpus, query("rotterdam harbour ")))).toEqual(["text", "pdf"]);
    expect(ids(search(corpus, query("rotterdam -draft ")))).toEqual(["pdf", "deck"]);
    expect(ids(search(corpus, query("title:harbour ")))).toEqual(["text"]);
    expect(ids(search(corpus, query("rotterdam kind:presentation ")))).toEqual(["deck"]);
    expect(ids(search(corpus, query("rotterdam -format:pdf ")))).toEqual(["deck", "text"]);
    expect(ids(search(corpus, query("format:pdf")))).toEqual(["pdf"]);
    expect(ids(search(corpus, query("kind:pres")))).toEqual(["deck"]);
  });

  it("matches a phrase only with its tokens adjacent, and never across a line break", () => {
    const corpus = buildIndex([
      doc("adjacent", "", "the Rotterdam figures rose", 1),
      doc("apart", "", "figures from Rotterdam", 2),
      doc("broken", "", "ends in Rotterdam\nfigures start here", 3),
    ]);
    const hits = search(corpus, query('"rotterdam figures"'));
    expect(ids(hits)).toEqual(["adjacent"]);
    expect(hits[0].positions).toEqual([4, 14]);
    // Both hold "figures" once; the shorter body ranks first.
    expect(ids(search(corpus, query('figures -"rotterdam figures"')))).toEqual(["apart", "broken"]);
  });

  it("matches the last word as a prefix while typing, and exactly after a space", () => {
    const corpus = buildIndex([doc("a", "", "Rotterdam port", 1), doc("b", "", "rotter", 2)]);
    expect(ids(search(corpus, query("rotter")))).toEqual(["b", "a"]);
    expect(ids(search(corpus, query("rotter ")))).toEqual(["b"]);
  });

  it("finds CJK text through its bigrams", () => {
    const corpus = buildIndex([doc("ja", "", "東京都庁の資料", 1), doc("zh", "", "京都大学", 2)]);
    expect(ids(search(corpus, query("東京都 ")))).toEqual(["ja"]);
    expect(ids(search(corpus, query("京都 ")))).toEqual(["zh", "ja"]);
    expect(ids(search(corpus, query("都庁の ")))).toEqual(["ja"]);
    expect(ids(search(corpus, query("大阪 ")))).toEqual([]);
    // A single character matches the bigrams it starts.
    expect(ids(search(corpus, query("都 ")))).toEqual(["zh", "ja"]);
    expect(ids(search(corpus, query("大 ")))).toEqual(["zh"]);
  });

  it("narrows hits with a filter without changing the survivors' scores", () => {
    const all = search(index, query("figures "));
    const narrowed = search(index, query("figures "), { filter: (id) => id === "d1" });
    expect(ids(narrowed)).toEqual(["d1"]);
    expect(narrowed[0].score).toBe(all.find((hit) => hit.id === "d1")?.score);
  });

  it("returns nothing for an empty query, and caps the list with limit", () => {
    expect(search(index, query(""))).toEqual([]);
    expect(search(index, query("figures "), { limit: 1 })).toHaveLength(1);
  });

  it("reports body positions in order, every token of every match", () => {
    const corpus = buildIndex([doc("a", "", "alpha beta alpha gamma", 1)]);
    expect(search(corpus, query("alpha gamma "))[0].positions).toEqual([0, 11, 17]);
  });
});

describe("the index is plain JSON, merged and replaced per document", () => {
  it("round-trips through JSON unchanged", () => {
    const index = buildIndex([doc("a", "Title", "Body text, café 東京", 1, { kind: ["text"] })]);
    expect(JSON.parse(JSON.stringify(index))).toEqual(index);
    expect(index.version).toBe(INDEX_VERSION);
  });

  it("stores terms sorted, so a hostile term is just a term", () => {
    const posting = indexDocument(doc("a", "", "constructor toString zebra apple valueOf", 1));
    expect(posting.body.terms).toEqual(["apple", "constructor", "tostring", "valueof", "zebra"]);
    expect(ids(search({ version: INDEX_VERSION, docs: [posting] }, query("constructor ")))).toEqual(["a"]);
    expect(ids(search({ version: INDEX_VERSION, docs: [posting] }, query("hasOwnProperty ")))).toEqual([]);
  });

  it("merges records, a later one replacing an earlier one with the same id", () => {
    const first = buildIndex([doc("a", "", "old words", 1), doc("b", "", "other", 1)]);
    const second = buildIndex([doc("a", "", "new words", 2)]);
    const merged = mergeIndex(first, second);
    expect(merged.docs.map((entry) => [entry.id, entry.lastModified])).toEqual([["a", 2], ["b", 1]]);
    expect(ids(search(merged, query("old ")))).toEqual([]);
    expect(ids(search(merged, query("new ")))).toEqual(["a"]);
    expect(mergeIndex(first, { version: INDEX_VERSION + 1, docs: second.docs }).docs[0].lastModified).toBe(1);
    expect(removeFromIndex(merged, ["a"]).docs.map((entry) => entry.id)).toEqual(["b"]);
  });
});

/** n distinct filler words. */
function filler(n: number): string {
  return Array.from({ length: n }, (_, i) => `word${i}`).join(" ");
}

describe("snippet — windows around a match", () => {

  it("at the start of a text: no leading ellipsis, the match marked", () => {
    const body = `Rotterdam figures ${filler(40)}`;
    const [hit] = search(buildIndex([doc("a", "", body)]), query("rotterdam "));
    const segments = snippet(body, hit, { radius: 30 });
    expect(segments[0]).toEqual({ text: "Rotterdam", match: true });
    expect(segments[segments.length - 1].text.endsWith("…")).toBe(true);
  });

  it("in the middle: both ends cut on word boundaries", () => {
    const body = `${filler(30)} the Rotterdam figures rose ${filler(30)}`;
    const [hit] = search(buildIndex([doc("a", "", body)]), query('"rotterdam figures"'));
    const segments = snippet(body, hit, { radius: 25 });
    const text = segments.map((segment) => segment.text).join("");
    expect(text.startsWith("…")).toBe(true);
    expect(text.endsWith("…")).toBe(true);
    expect(segments.filter((segment) => segment.match).map((segment) => segment.text)).toEqual(["Rotterdam figures"]);
    // Whole words only at both cuts.
    expect(text).toMatch(/^…word\d+ /);
    expect(text).toMatch(/ word\d+…$/);
  });

  it("at the end of a text: no trailing ellipsis", () => {
    const body = `${filler(40)} final Rotterdam`;
    const [hit] = search(buildIndex([doc("a", "", body)]), query("rotterdam "));
    const segments = snippet(body, hit, { radius: 30 });
    expect(segments[segments.length - 1]).toEqual({ text: "Rotterdam", match: true });
    expect(segments[0].text.startsWith("…")).toBe(true);
  });

  it("collapses line breaks, and a title-only hit has no snippet", () => {
    const body = "first block\nsecond Rotterdam block";
    const [hit] = search(buildIndex([doc("a", "", body)]), query("rotterdam "));
    expect(snippet(body, hit).map((segment) => segment.text).join("")).toBe("first block second Rotterdam block");
    expect(snippet("", { positions: [] })).toEqual([]);
  });

  it("gives at most MAX_SNIPPETS_PER_DOC non-overlapping snippets", () => {
    const body = Array.from({ length: 8 }, (_, i) => `Rotterdam ${filler(30)} ${i}`).join(" ");
    const [hit] = search(buildIndex([doc("a", "", body)]), query("rotterdam "));
    expect(MAX_SNIPPETS_PER_DOC).toBe(3);
    expect(snippets(body, hit)).toHaveLength(3);
  });
});

describe("result helpers", () => {
  it("marks the matched words of a title", () => {
    expect(titleSegments("Rotterdam Harbour report", query("harb"))).toEqual([
      { text: "Rotterdam ", match: false },
      { text: "Harbour", match: true },
      { text: " report", match: false },
    ]);
  });

  it("hands a find bar the clause whose match comes first, as typed", () => {
    const body = "Figures first. Then Rotterdam.";
    const parsed = query('rotterdam "figures first"');
    const [hit] = search(buildIndex([doc("a", "", body)]), parsed);
    expect(findTextFor(parsed, hit, body)).toBe("figures first");
    expect(findTextFor(query("title:x rotter"), { positions: [] }, "")).toBe("rotter");
    expect(findTextFor(query("kind:text"), { positions: [] }, "")).toBe("");
  });
});
