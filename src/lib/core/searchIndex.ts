// Full-text search over the library (SPEC-114 — the normative source for the
// tokeniser, the query grammar and the ranking function). Product-neutral and
// pure: a document reaches this module as `{ id, lastModified, title, body }`
// plus optional facets (`kind`, `format`, … — values it compares, never
// interprets), and nothing here knows a document model, a deck or a PDF.
// No DOM, no framework, no clock, no randomness: the same inputs always give
// the same index, the same hits in the same order and the same snippets.
//
// The index is plain JSON (the store's standing requirement): one
// `DocPostings` per document — its own posting lists, which is what one
// persisted index record holds — and a `SearchIndex` is just a list of them,
// so the searchable index is the merge of the loaded records and one
// document re-indexing replaces one entry. Terms are stored as a sorted
// array beside their posting lists rather than as object keys: a term is
// arbitrary user text ("__proto__" included), and a sorted array also answers
// the prefix lookups the last typed word needs with two binary searches.
//
// The tokeniser is part of the stored format: tokens are what the records
// hold, so ANY change to its rules (or to scoring) must bump INDEX_VERSION,
// which invalidates every stored record. Each rule has its own test.

/** Bumped by ANY tokeniser or scoring change — it invalidates every stored
 * index record. */
export const INDEX_VERSION = 1;

/** BM25's term-frequency saturation. */
export const BM25_K1 = 1.2;
/** BM25's length normalisation. */
export const BM25_B = 0.75;
/** The title field's weight against the body's. */
export const TITLE_BOOST = 3;
/** Tokens longer than this are cut to it (in UTF-16 code units, never
 * splitting a surrogate pair). */
export const MAX_TOKEN_LENGTH = 64;
/** How many snippets one result shows at most. */
export const MAX_SNIPPETS_PER_DOC = 3;
/** Characters of context a snippet keeps either side of its match. */
export const SNIPPET_RADIUS = 60;

// --- the tokeniser -------------------------------------------------------------

/** One token: its normalised text and where it came from in the ORIGINAL
 * string — `start` is the offset of its first character, `end` one past its
 * last, so a snippet can mark exactly the source text a match covers. */
export interface Token {
  token: string;
  start: number;
  end: number;
}

/** How the tokeniser treats one normalised character. */
const CharClass = {
  Separator: 0,
  Word: 1,
  Cjk: 2,
  /** `'` or `-`: part of a token only between two word characters. */
  Joiner: 3,
  /** `.` or `,`: part of a token only between two digits (numbers whole). */
  NumberJoiner: 4,
} as const;
type CharClass = (typeof CharClass)[keyof typeof CharClass];

const WORD_CHAR = /^[\p{L}\p{N}]$/u;
const DIGIT_CHAR = /^\p{N}$/u;
const CJK_CHAR = /^[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\u30fc]$/u;
const COMBINING_MARK = /\p{M}/gu;

/** Typographic variants folded to the two joiners the rules name, so a
 * curly apostrophe typed by a word processor matches a straight one. */
const JOINER_FOLDS: Record<string, string> = {
  "\u2019": "'",
  "\u2018": "'",
  "\u02bc": "'",
  "\u2010": "-",
  "\u2011": "-",
};

/** One code point's normalised form: NFKD-decomposed, combining marks
 * stripped, lowercased. May be "" (a lone combining mark) or several characters
 * (a ligature, a vulgar fraction). Cached — real text reuses few characters. */
const foldCache = new Map<string, string>();

function fold(char: string): string {
  let folded = foldCache.get(char);
  if (folded === undefined) {
    const joiner = JOINER_FOLDS[char];
    // Lowercased on both sides of the decomposition: `İ` lowercases to
    // `i` + a combining dot (stripped), and a compatibility form such as
    // a mathematical bold `𝐀` only decomposes to an uppercase `A`.
    folded = joiner ?? char.toLowerCase().normalize("NFKD").replace(COMBINING_MARK, "").toLowerCase();
    foldCache.set(char, folded);
  }
  return folded;
}

const classCache = new Map<string, CharClass>();

function classify(char: string): CharClass {
  let cls = classCache.get(char);
  if (cls === undefined) {
    if (char === "'" || char === "-") {
      cls = CharClass.Joiner;
    } else if (char === "." || char === ",") {
      cls = CharClass.NumberJoiner;
    } else if (CJK_CHAR.test(char)) {
      cls = CharClass.Cjk;
    } else if (WORD_CHAR.test(char)) {
      cls = CharClass.Word;
    } else {
      cls = CharClass.Separator;
    }
    classCache.set(char, cls);
  }
  return cls;
}

function isDigit(char: string): boolean {
  return DIGIT_CHAR.test(char);
}

/** Cuts a token to MAX_TOKEN_LENGTH code units without splitting a pair. */
function capToken(token: string): string {
  if (token.length <= MAX_TOKEN_LENGTH) {
    return token;
  }
  const last = token.charCodeAt(MAX_TOKEN_LENGTH - 1);
  return token.slice(0, last >= 0xd800 && last <= 0xdbff ? MAX_TOKEN_LENGTH - 1 : MAX_TOKEN_LENGTH);
}

/** One normalised character and the source range it came from. */
interface FoldedChar {
  char: string;
  cls: CharClass;
  start: number;
  end: number;
}

/** The normalised character stream of `text`, each character carrying the
 * source code point's range (a ligature's two letters share one range). */
function foldedChars(text: string): FoldedChar[] {
  const out: FoldedChar[] = [];
  for (let i = 0; i < text.length; ) {
    const code = text.charCodeAt(i);
    const width = code >= 0xd800 && code <= 0xdbff && i + 1 < text.length ? 2 : 1;
    const end = i + width;
    if (code < 0x80) {
      // ASCII fast path: no decomposition, one-to-one lowercase.
      const char = code >= 0x41 && code <= 0x5a ? String.fromCharCode(code + 32) : text[i];
      out.push({ char, cls: classify(char), start: i, end });
    } else {
      for (const char of fold(text.slice(i, end))) {
        out.push({ char, cls: classify(char), start: i, end });
      }
    }
    i = end;
  }
  return out;
}

/** Emits a run of CJK/Kana characters as overlapping bigrams (a lone
 * character as itself): a language written without spaces is searchable at
 * all only this way. */
function pushCjkRun(run: FoldedChar[], tokens: Token[]): void {
  if (run.length === 1) {
    tokens.push({ token: run[0].char, start: run[0].start, end: run[0].end });
    return;
  }
  for (let i = 0; i + 1 < run.length; i++) {
    tokens.push({ token: run[i].char + run[i + 1].char, start: run[i].start, end: run[i + 1].end });
  }
}

/** Splits text into search tokens, by rule:
 * - lowercase, Unicode NFKD with combining marks stripped (`café` → `cafe`);
 * - a token is a run of Unicode letters and digits; anything else splits,
 *   except `'` or `-` with a letter or digit on BOTH sides (`don't`,
 *   `e-mail`) — curly apostrophes and Unicode hyphens fold to those two;
 * - numbers are kept whole: `.` or `,` between two digits stays inside the
 *   token (`3.14`, `1,250`), so a figure is one token;
 * - a token is cut to MAX_TOKEN_LENGTH;
 * - a run of CJK ideographs or Kana is indexed as overlapping bigrams.
 * Offsets are into the original `text`. */
export function tokenise(text: string): Token[] {
  const chars = foldedChars(text);
  const tokens: Token[] = [];
  let word = "";
  let wordStart = 0;
  let wordEnd = 0;
  let cjk: FoldedChar[] = [];
  const endWord = (): void => {
    if (word !== "") {
      tokens.push({ token: capToken(word), start: wordStart, end: wordEnd });
      word = "";
    }
  };
  const endCjk = (): void => {
    if (cjk.length > 0) {
      pushCjkRun(cjk, tokens);
      cjk = [];
    }
  };
  for (let i = 0; i < chars.length; i++) {
    const current = chars[i];
    if (current.cls === CharClass.Word) {
      endCjk();
      if (word === "") {
        wordStart = current.start;
      }
      word += current.char;
      wordEnd = current.end;
    } else if (current.cls === CharClass.Cjk) {
      endWord();
      cjk.push(current);
    } else if (current.cls === CharClass.Joiner || current.cls === CharClass.NumberJoiner) {
      const previous = word === "" ? undefined : word[word.length - 1];
      const next = chars[i + 1];
      const joins =
        previous !== undefined &&
        next !== undefined &&
        next.cls === CharClass.Word &&
        (current.cls === CharClass.Joiner || (isDigit(previous) && isDigit(next.char)));
      if (joins) {
        word += current.char;
      } else {
        endWord();
        endCjk();
      }
    } else {
      endWord();
      endCjk();
    }
  }
  endWord();
  endCjk();
  return tokens;
}

/** Whether a token came from a CJK/Kana run (a bigram or a lone character):
 * such a token, typed on its own, also matches the bigrams it starts. */
function isCjkToken(token: string): boolean {
  return token.length > 0 && classify(String.fromCodePoint(token.codePointAt(0) ?? 0)) === CharClass.Cjk;
}

// --- the index -------------------------------------------------------------------

/** What a document hands the index. `facets` are compared, never
 * interpreted: `kind:`/`format:` filters match a value listed here. */
export interface IndexedDoc {
  id: string;
  lastModified: number;
  title: string;
  body: string;
  /** Lowercase values per facet, e.g. `{ kind: ["text"], format: ["docx"] }`. */
  facets?: Record<string, string[]>;
}

/** One field's posting lists: `terms` sorted, `positions[i]` the ordinals
 * (token positions, ascending) at which `terms[i]` occurs. */
export interface FieldPostings {
  terms: string[];
  positions: number[][];
  /** How many tokens the field holds (BM25's document length). */
  length: number;
}

/** One document's own posting lists — what a persisted index record holds. */
export interface DocPostings {
  id: string;
  lastModified: number;
  facets: Record<string, string[]>;
  title: FieldPostings;
  body: FieldPostings;
  /** Body ordinal → character offset of that token in the body text (-1
   * for an ordinal a line break skipped). */
  starts: number[];
}

/** The searchable index: plain JSON, one entry per document. */
export interface SearchIndex {
  version: number;
  docs: DocPostings[];
}

/** Each token's ordinal: consecutive, except that a line break between two
 * tokens leaves one ordinal empty — the extractors put a line break between
 * blocks, so a phrase never matches across a paragraph (where no find bar
 * could land on it). */
function ordinalsOf(text: string, tokens: Token[]): number[] {
  const ordinals: number[] = [];
  let ordinal = -1;
  let previousEnd = 0;
  for (const token of tokens) {
    const gap = ordinal >= 0 && text.slice(previousEnd, token.start).includes("\n");
    ordinal += gap ? 2 : 1;
    ordinals.push(ordinal);
    previousEnd = token.end;
  }
  return ordinals;
}

function fieldPostings(tokens: Token[], ordinals: number[]): FieldPostings {
  const byTerm = new Map<string, number[]>();
  tokens.forEach((token, i) => {
    const list = byTerm.get(token.token);
    if (list) {
      list.push(ordinals[i]);
    } else {
      byTerm.set(token.token, [ordinals[i]]);
    }
  });
  const terms = [...byTerm.keys()].toSorted(compareStrings);
  return { terms, positions: terms.map((term) => byTerm.get(term) as number[]), length: tokens.length };
}

function compareStrings(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}

/** One document's postings. */
export function indexDocument(doc: IndexedDoc): DocPostings {
  const titleTokens = tokenise(doc.title);
  const bodyTokens = tokenise(doc.body);
  const bodyOrdinals = ordinalsOf(doc.body, bodyTokens);
  // Ordinal → offset; an ordinal a line break skipped holds -1.
  const starts: number[] = [];
  bodyTokens.forEach((token, i) => {
    while (starts.length < bodyOrdinals[i]) {
      starts.push(-1);
    }
    starts.push(token.start);
  });
  return {
    id: doc.id,
    lastModified: doc.lastModified,
    facets: normaliseFacets(doc.facets),
    title: fieldPostings(titleTokens, ordinalsOf(doc.title, titleTokens)),
    body: fieldPostings(bodyTokens, bodyOrdinals),
    starts,
  };
}

function normaliseFacets(facets: Record<string, string[]> | undefined): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  for (const [facet, values] of Object.entries(facets ?? {})) {
    out[facet.toLowerCase()] = values.map((value) => value.toLowerCase());
  }
  return out;
}

export function buildIndex(docs: IndexedDoc[]): SearchIndex {
  return { version: INDEX_VERSION, docs: docs.map(indexDocument) };
}

/** Merges indexes into one: every document of every part, a later part's
 * entry replacing an earlier one with the same id (a re-indexed document).
 * Parts built under another INDEX_VERSION are skipped — their tokens were
 * cut by other rules. Order: first appearance. */
export function mergeIndex(...parts: SearchIndex[]): SearchIndex {
  const byId = new Map<string, DocPostings>();
  for (const part of parts) {
    if (part.version !== INDEX_VERSION) {
      continue;
    }
    for (const doc of part.docs) {
      byId.set(doc.id, doc);
    }
  }
  return { version: INDEX_VERSION, docs: [...byId.values()] };
}

/** The index without the given documents. */
export function removeFromIndex(index: SearchIndex, ids: Iterable<string>): SearchIndex {
  const drop = new Set(ids);
  return { version: index.version, docs: index.docs.filter((doc) => !drop.has(doc.id)) };
}

// --- the query grammar -------------------------------------------------------------

/** A word or phrase to match. A word the tokeniser splits into several
 * tokens (`U.S.`, a CJK run) matches like a phrase: its tokens adjacent. */
export interface QueryTerm {
  tokens: string[];
  /** The clause as typed, quotes removed — what a find bar should look for. */
  text: string;
  phrase: boolean;
  /** The last token matches as a prefix: the last bare word while typing. */
  prefix: boolean;
  exclude: boolean;
  /** `title:` — match the title only. */
  field?: "title";
  /** The clause's span in the input. */
  start: number;
  end: number;
}

/** A `kind:` or `format:` filter. */
export interface QueryFacet {
  facet: "kind" | "format";
  value: string;
  prefix: boolean;
  exclude: boolean;
  start: number;
  end: number;
}

export interface Query {
  terms: QueryTerm[];
  facets: QueryFacet[];
}

/** Why a query could not be parsed, and where (an offset into the input). */
export interface QueryError {
  error: string;
  position: number;
}

export function isSearchQueryError(value: Query | QueryError): value is QueryError {
  return "error" in value;
}

const FILTER = /^(title|kind|format):/i;

function isSpace(char: string | undefined): boolean {
  return char !== undefined && /\s/.test(char);
}

/** Parses what the user typed. Total: every input is a Query or a
 * QueryError naming the offending position — it never throws, and nothing
 * unparseable is ever read as "match everything".
 * - space-separated clauses are ANDed;
 * - `"a phrase"` matches its tokens adjacent;
 * - a leading `-` excludes (a word, a phrase or a filter);
 * - `title:`, `kind:` and `format:` filter (case-insensitive names);
 * - the LAST clause, when it is a bare word or filter value and the input
 *   does not end in whitespace, is a prefix match — the user is still
 *   typing it; an explicit trailing space ends that. Exclusions never are.
 * An empty (or all-whitespace) input is an empty query, which matches
 * nothing. */
export function parseQuery(input: string): Query | QueryError {
  const terms: QueryTerm[] = [];
  const facets: QueryFacet[] = [];
  let lastBare: { clause: QueryTerm | QueryFacet } | null = null;
  let pos = 0;
  const len = input.length;
  while (pos < len) {
    while (pos < len && isSpace(input[pos])) {
      pos++;
    }
    if (pos >= len) {
      break;
    }
    const clauseStart = pos;
    let exclude = false;
    if (input[pos] === "-") {
      exclude = true;
      pos++;
      if (pos >= len || isSpace(input[pos])) {
        return { error: "A “-” needs a word or phrase right after it", position: clauseStart };
      }
      if (input[pos] === "-") {
        return { error: "Only one “-” per word", position: pos };
      }
    }
    let field: "title" | "kind" | "format" | undefined;
    const filter = FILTER.exec(input.slice(pos, pos + 7));
    if (filter) {
      field = filter[1].toLowerCase() as "title" | "kind" | "format";
      pos += filter[0].length;
      if (pos >= len || isSpace(input[pos])) {
        return { error: `“${field}:” needs a value right after it`, position: pos };
      }
      if (input[pos] === "-") {
        return { error: `Put the “-” before the filter: -${field}:…`, position: pos };
      }
    }
    let text: string;
    let phrase = false;
    let valueStart = pos;
    if (input[pos] === '"') {
      const close = input.indexOf('"', pos + 1);
      if (close === -1) {
        return { error: "This quote is never closed", position: pos };
      }
      text = input.slice(pos + 1, close);
      if (text.trim() === "") {
        return { error: "This phrase is empty", position: pos };
      }
      valueStart = pos + 1;
      phrase = true;
      pos = close + 1;
      if (pos < len && !isSpace(input[pos])) {
        return { error: "Leave a space after the closing quote", position: pos };
      }
    } else {
      let end = pos;
      while (end < len && !isSpace(input[end])) {
        if (input[end] === '"') {
          return { error: "A quote can only start a phrase", position: end };
        }
        end++;
      }
      text = input.slice(pos, end);
      pos = end;
    }
    if (field === "kind" || field === "format") {
      const value = text.trim().toLowerCase();
      const facet: QueryFacet = { facet: field, value, prefix: false, exclude, start: clauseStart, end: pos };
      facets.push(facet);
      lastBare = phrase ? null : { clause: facet };
      continue;
    }
    const tokens = tokenise(text).map((token) => token.token);
    if (tokens.length === 0) {
      return { error: `“${text}” has no letters or digits to search for`, position: valueStart };
    }
    const term: QueryTerm = {
      tokens,
      text,
      phrase: phrase || tokens.length > 1,
      prefix: false,
      exclude,
      ...(field === "title" ? { field: "title" as const } : {}),
      start: clauseStart,
      end: pos,
    };
    terms.push(term);
    lastBare = phrase ? null : { clause: term };
  }
  // The last bare clause is still being typed unless whitespace ended it.
  if (lastBare && !lastBare.clause.exclude && !isSpace(input[len - 1]) && lastBare.clause.end === len) {
    lastBare.clause.prefix = true;
  }
  const positive = terms.some((term) => !term.exclude) || facets.some((facet) => !facet.exclude);
  const firstExclusion = [...terms, ...facets].filter((clause) => clause.exclude).toSorted((a, b) => a.start - b.start)[0];
  if (!positive && firstExclusion) {
    return { error: "Add something to search for — “-” only narrows a search", position: firstExclusion.start };
  }
  return { terms, facets };
}

// --- matching ------------------------------------------------------------------------

/** Index of the first element of sorted `terms` that is >= `key`. */
function lowerBound(terms: string[], key: string): number {
  let lo = 0;
  let hi = terms.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (terms[mid] < key) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** The posting lists one query token selects in a field: the exact term, or
 * every term it prefixes. A lone CJK character (which only a one-character
 * run indexes as itself) also selects the bigrams it starts. */
function postingsFor(field: FieldPostings, token: string, prefix: boolean): number[][] {
  const lists: number[][] = [];
  const from = lowerBound(field.terms, token);
  if (prefix || (isCjkToken(token) && [...token].length === 1)) {
    for (let i = from; i < field.terms.length && field.terms[i].startsWith(token); i++) {
      lists.push(field.positions[i]);
    }
  } else if (from < field.terms.length && field.terms[from] === token) {
    lists.push(field.positions[from]);
  }
  return lists;
}

/** The ordinals at which a clause's tokens all occur in order, adjacent —
 * each a run start. A single-token clause's ordinals are its postings. */
function clauseOccurrences(field: FieldPostings, term: QueryTerm): number[] {
  const last = term.tokens.length - 1;
  const perToken = term.tokens.map((token, i) => mergeSorted(postingsFor(field, token, term.prefix && i === last)));
  if (perToken.some((list) => list.length === 0)) {
    return [];
  }
  let runs = perToken[0];
  for (let i = 1; i < perToken.length; i++) {
    runs = adjacent(runs, perToken[i], i);
  }
  return runs;
}

/** Starts in `starts` that have `offset`-later ordinals in `next`. Both
 * ascending: a linear merge. */
function adjacent(starts: number[], next: number[], offset: number): number[] {
  const out: number[] = [];
  let j = 0;
  for (const start of starts) {
    const want = start + offset;
    while (j < next.length && next[j] < want) {
      j++;
    }
    if (j < next.length && next[j] === want) {
      out.push(start);
    }
  }
  return out;
}

/** One ascending list from several ascending lists. */
function mergeSorted(lists: number[][]): number[] {
  if (lists.length === 0) {
    return [];
  }
  if (lists.length === 1) {
    return lists[0];
  }
  return Array.from(Float64Array.from(lists.flat()).toSorted());
}

/** The `k` smallest values of several ascending lists, ascending: a k-way
 * merge over a binary heap of list heads, so a prefix that selects
 * hundreds of terms costs k steps, not a sort of every occurrence. */
function smallestOf(lists: number[][], k: number): number[] {
  if (!Number.isFinite(k)) {
    return mergeSorted(lists);
  }
  if (lists.length <= 1) {
    return (lists[0] ?? []).slice(0, k);
  }
  // Heap entries: [value, list, index in list].
  const heap: Array<[number, number, number]> = [];
  const less = (i: number, j: number): boolean => heap[i][0] < heap[j][0];
  const swap = (i: number, j: number): void => {
    [heap[i], heap[j]] = [heap[j], heap[i]];
  };
  const push = (entry: [number, number, number]): void => {
    heap.push(entry);
    for (let i = heap.length - 1; i > 0 && less(i, (i - 1) >> 1); i = (i - 1) >> 1) {
      swap(i, (i - 1) >> 1);
    }
  };
  const pop = (): [number, number, number] => {
    const top = heap[0];
    const last = heap.pop() as [number, number, number];
    if (heap.length > 0) {
      heap[0] = last;
      for (let i = 0; ; ) {
        const left = i * 2 + 1;
        const smallest = left + 1 < heap.length && less(left + 1, left) ? left + 1 : left;
        if (smallest >= heap.length || !less(smallest, i)) break;
        swap(i, smallest);
        i = smallest;
      }
    }
    return top;
  };
  lists.forEach((list, i) => {
    if (list.length > 0) push([list[0], i, 0]);
  });
  const out: number[] = [];
  while (out.length < k && heap.length > 0) {
    const [value, list, at] = pop();
    out.push(value);
    if (at + 1 < lists[list].length) push([lists[list][at + 1], list, at + 1]);
  }
  return out;
}

/** How many times a clause occurs in a field — what BM25 needs, without
 * merging a single-token clause's posting lists. */
function clauseCount(field: FieldPostings, term: QueryTerm): number {
  if (term.tokens.length === 1) {
    let count = 0;
    for (const list of postingsFor(field, term.tokens[0], term.prefix)) {
      count += list.length;
    }
    return count;
  }
  return clauseOccurrences(field, term).length;
}

/** The ordinals at which a clause's first `k` occurrences start. */
function earliestOccurrences(field: FieldPostings, term: QueryTerm, k: number): number[] {
  if (term.tokens.length === 1) {
    return smallestOf(postingsFor(field, term.tokens[0], term.prefix), k);
  }
  const all = clauseOccurrences(field, term);
  return Number.isFinite(k) ? all.slice(0, k) : all;
}

function facetMatches(doc: DocPostings, facet: QueryFacet): boolean {
  const values = doc.facets[facet.facet] ?? [];
  return values.some((value) => (facet.prefix ? value.startsWith(facet.value) : value === facet.value));
}

// --- ranking ---------------------------------------------------------------------------

/** One result. `positions` are character offsets into the body, ascending:
 * every token of every positive match, for snippets and "open at match". */
export interface SearchHit {
  id: string;
  score: number;
  positions: number[];
  /** Matched occurrences in the title. */
  titleMatches: number;
  /** Matched occurrences in the body (a phrase counts once per occurrence). */
  bodyMatches: number;
  lastModified: number;
}

/** BM25's inverse document frequency, the non-negative form:
 * ln(1 + (N − df + 0.5) / (df + 0.5)). */
export function bm25Idf(docCount: number, docFrequency: number): number {
  return Math.log(1 + (docCount - docFrequency + 0.5) / (docFrequency + 0.5));
}

/** One field's BM25 contribution for one term:
 * idf × tf × (k1 + 1) / (tf + k1 × (1 − b + b × dl / avgdl)). */
export function bm25Term(tf: number, idf: number, length: number, averageLength: number): number {
  if (tf === 0) {
    return 0;
  }
  const norm = averageLength > 0 ? length / averageLength : 1;
  return (idf * tf * (BM25_K1 + 1)) / (tf + BM25_K1 * (1 - BM25_B + BM25_B * norm));
}

/** Per document, per positive clause: occurrences in the body and title. */
interface ClauseMatch {
  body: number;
  title: number;
}

/** Finds and ranks documents. Every positive clause must match (in the body
 * or the title; a `title:` clause in the title), no excluded clause may, and
 * every facet filter must hold. Score = Σ over positive clauses of
 * BM25(body) + TITLE_BOOST × BM25(title), with k1 = 1.2, b = 0.75, the
 * non-negative idf over the whole index (a phrase or prefix clause is one
 * term: its occurrences are its tf, the documents holding one its df).
 * Ties: newest `lastModified` first, then id. `limit` caps how many hits are
 * returned; `filter` keeps only the documents it accepts as hits, while
 * every document still counts towards N, df and the average lengths —
 * narrowing a search never changes a survivor's score. */
export function search(
  index: SearchIndex,
  query: Query,
  options: { limit?: number; filter?: (id: string) => boolean } = {},
): SearchHit[] {
  const ranked = rank(index, query, options.filter);
  return withPositions(index, query, options.limit === undefined ? ranked : ranked.slice(0, Math.max(0, options.limit)));
}

/** Every hit, scored and ordered as `search` orders them, without body
 * positions — the cheap part, so a caller can count all hits and then
 * fill positions (withPositions) for only the ones it shows. */
export function rank(index: SearchIndex, query: Query, filter?: (id: string) => boolean): SearchHit[] {
  const positive = query.terms.filter((term) => !term.exclude);
  const excluded = query.terms.filter((term) => term.exclude);
  if (positive.length === 0 && !query.facets.some((facet) => !facet.exclude)) {
    return [];
  }
  const docs = index.docs;
  const docCount = docs.length;
  let bodyTotal = 0;
  let titleTotal = 0;
  for (const doc of docs) {
    bodyTotal += doc.body.length;
    titleTotal += doc.title.length;
  }
  const avgBody = docCount > 0 ? bodyTotal / docCount : 0;
  const avgTitle = docCount > 0 ? titleTotal / docCount : 0;

  // Every clause is evaluated over every document: df is a property of the
  // whole index, not of the documents that survive the other clauses.
  const matches: ClauseMatch[][] = positive.map((term) =>
    docs.map((doc) => ({
      body: term.field === "title" ? 0 : clauseCount(doc.body, term),
      title: clauseCount(doc.title, term),
    })),
  );
  const bodyDf = matches.map((perDoc) => perDoc.filter((match) => match.body > 0).length);
  const titleDf = matches.map((perDoc) => perDoc.filter((match) => match.title > 0).length);

  const hits: SearchHit[] = [];
  docs.forEach((doc, d) => {
    if (filter && !filter(doc.id)) {
      return;
    }
    if (!query.facets.every((facet) => facetMatches(doc, facet) !== facet.exclude)) {
      return;
    }
    if (!positive.every((_, t) => matches[t][d].body > 0 || matches[t][d].title > 0)) {
      return;
    }
    if (excluded.some((term) => clauseCount(doc.title, term) > 0 || (term.field !== "title" && clauseCount(doc.body, term) > 0))) {
      return;
    }
    let score = 0;
    let titleMatches = 0;
    let bodyMatches = 0;
    positive.forEach((_, t) => {
      const match = matches[t][d];
      score += bm25Term(match.body, bm25Idf(docCount, bodyDf[t]), doc.body.length, avgBody);
      score += TITLE_BOOST * bm25Term(match.title, bm25Idf(docCount, titleDf[t]), doc.title.length, avgTitle);
      titleMatches += match.title;
      bodyMatches += match.body;
    });
    hits.push({ id: doc.id, score, positions: [], titleMatches, bodyMatches, lastModified: doc.lastModified });
  });
  return hits.toSorted((a, b) => b.score - a.score || b.lastModified - a.lastModified || compareStrings(a.id, b.id));
}

/** The hits with their body positions filled in (see SearchHit). With
 * `maxPositions`, only each hit's earliest positions — all a snippet or
 * "open at match" reads — so a one-letter prefix stays cheap. */
export function withPositions(index: SearchIndex, query: Query, hits: SearchHit[], maxPositions = Infinity): SearchHit[] {
  const positive = query.terms.filter((term) => !term.exclude && term.field !== "title");
  const byId = new Map(index.docs.map((doc) => [doc.id, doc]));
  return hits.map((hit) => {
    const doc = byId.get(hit.id);
    return doc ? { ...hit, positions: bodyPositions(doc, positive, maxPositions) } : hit;
  });
}

/** The body character offsets a document's positive matches cover — each
 * token of each occurrence — ascending and once each; the first `max`. */
function bodyPositions(doc: DocPostings, positive: QueryTerm[], max: number): number[] {
  const offsets = new Set<number>();
  for (const term of positive) {
    for (const ordinal of earliestOccurrences(doc.body, term, max)) {
      for (let k = 0; k < term.tokens.length; k++) {
        const start = doc.starts[ordinal + k];
        if (start !== undefined && start >= 0) {
          offsets.add(start);
        }
      }
    }
  }
  const sorted = Array.from(Float64Array.from(offsets).toSorted());
  return Number.isFinite(max) ? sorted.slice(0, max) : sorted;
}

// --- snippets ----------------------------------------------------------------------------

/** A run of snippet text, flagged as matched or not — the same shape as
 * SPEC-067's title highlight segments, so it renders the same way. */
export interface SnippetSegment {
  text: string;
  match: boolean;
}

function isWhitespaceCode(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13 || code === 12 || code === 0xa0;
}

/** Segments of `text`[from, to) with the given source ranges marked; runs of
 * whitespace collapse to one space (a body joins blocks with newlines). */
function markRanges(text: string, from: number, to: number, ranges: Array<[number, number]>): SnippetSegment[] {
  const segments: SnippetSegment[] = [];
  const push = (piece: string, match: boolean): void => {
    const clean = piece.replace(/\s+/g, " ");
    if (clean === "") {
      return;
    }
    const last = segments[segments.length - 1];
    if (last && last.match === match) {
      last.text += clean;
    } else {
      segments.push({ text: clean, match });
    }
  };
  let cursor = from;
  for (const [start, end] of ranges) {
    if (start > cursor) {
      push(text.slice(cursor, start), false);
    }
    push(text.slice(Math.max(start, cursor), end), true);
    cursor = Math.max(cursor, end);
  }
  if (cursor < to) {
    push(text.slice(cursor, to), false);
  }
  return segments;
}

/** The window around one match of `hit` in `body` (`options.match`, an index
 * into hit.positions, default the first), at most `radius` characters of
 * context either side, widened to whole words, with every matched token
 * inside it marked and "…" where the body was cut. A hit with no body
 * position (a title-only match) has no snippet. */
export function snippet(
  body: string,
  hit: Pick<SearchHit, "positions">,
  options: { radius?: number; match?: number } = {},
): SnippetSegment[] {
  const center = hit.positions[options.match ?? 0];
  if (center === undefined || center >= body.length) {
    return [];
  }
  const radius = options.radius ?? SNIPPET_RADIUS;
  let from = Math.max(0, center - radius);
  let to = Math.min(body.length, center + radius);
  // Widen to whole words: never start or end inside one.
  while (from > 0 && !isWhitespaceCode(body.charCodeAt(from - 1)) && center - from < radius * 2) {
    from--;
  }
  while (to < body.length && !isWhitespaceCode(body.charCodeAt(to)) && to - center < radius * 2) {
    to++;
  }
  const marked = new Set(hit.positions.filter((position) => position >= from && position < to));
  const ranges: Array<[number, number]> = [];
  for (const token of tokenise(body.slice(from, to))) {
    const start = from + token.start;
    if (!marked.has(start)) {
      continue;
    }
    const end = from + token.end;
    const last = ranges[ranges.length - 1];
    // Matches with only whitespace between them read as one highlight — a
    // phrase is marked whole, not word by word.
    if (last && (start <= last[1] || /^[^\S\n]+$/.test(body.slice(last[1], start)))) {
      last[1] = Math.max(last[1], end);
    } else {
      ranges.push([start, end]);
    }
  }
  const segments = markRanges(body, from, to, ranges);
  if (segments.length > 0 && from > 0) {
    segments[0] = segments[0].match ? segments[0] : { text: `…${segments[0].text.trimStart()}`, match: false };
    if (segments[0].match) {
      segments.unshift({ text: "…", match: false });
    }
  }
  const tail = segments[segments.length - 1];
  if (tail && to < body.length) {
    if (tail.match) {
      segments.push({ text: "…", match: false });
    } else {
      tail.text = `${tail.text.trimEnd()}…`;
    }
  }
  return segments;
}

/** Up to `max` snippets for a hit, each around the first match not already
 * shown by an earlier one, in body order. */
export function snippets(
  body: string,
  hit: Pick<SearchHit, "positions">,
  options: { radius?: number; max?: number } = {},
): SnippetSegment[][] {
  const radius = options.radius ?? SNIPPET_RADIUS;
  const max = options.max ?? MAX_SNIPPETS_PER_DOC;
  const out: SnippetSegment[][] = [];
  let shownUntil = -1;
  hit.positions.forEach((position, index) => {
    if (out.length >= max || position <= shownUntil) {
      return;
    }
    const segments = snippet(body, hit, { radius, match: index });
    if (segments.length > 0) {
      out.push(segments);
      shownUntil = position + radius;
    }
  });
  return out;
}

/** A title split into segments with the tokens a query's positive clauses
 * match marked — the result row's title highlight. */
export function titleSegments(title: string, query: Query): SnippetSegment[] {
  const wanted = query.terms.filter((term) => !term.exclude);
  const ranges: Array<[number, number]> = [];
  for (const token of tokenise(title)) {
    const hit = wanted.some((term) =>
      term.tokens.some((want, i) =>
        term.prefix && i === term.tokens.length - 1 ? token.token.startsWith(want) : token.token === want,
      ),
    );
    if (hit) {
      ranges.push([token.start, token.end]);
    }
  }
  return markRanges(title, 0, title.length, ranges);
}

/** What a find bar should be handed to land on a hit: the positive clause
 * whose first body occurrence comes first, as typed (a phrase without its
 * quotes); the first positive clause when none matched the body. "" for a
 * query of filters only. */
export function findTextFor(query: Query, hit: Pick<SearchHit, "positions">, body: string): string {
  const positive = query.terms.filter((term) => !term.exclude && term.field !== "title");
  if (positive.length === 0) {
    return "";
  }
  const first = hit.positions[0];
  if (first !== undefined) {
    const token = tokenise(body.slice(first, first + MAX_TOKEN_LENGTH * 2))[0];
    const found = token && positive.find((term) => term.tokens[0] !== undefined && (term.prefix && term.tokens.length === 1 ? token.token.startsWith(term.tokens[0]) : token.token === term.tokens[0]));
    if (found) {
      return found.text.trim().replace(/\s+/g, " ");
    }
  }
  return positive[0].text.trim().replace(/\s+/g, " ");
}
