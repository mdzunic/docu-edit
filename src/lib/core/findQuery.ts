// The find query, product-neutral (SPEC-098's move, performed by SPEC-134):
// compiling what the user typed in the find bar under its options into a
// RegExp. Moved unchanged, names included, out of src/lib/findReplace.ts
// (SPEC-017), which re-exports every name here, so the text editor's imports
// and findReplace.spec.ts are untouched — and the deck and the sheet search
// (SPEC-098, SPEC-134) compile their queries without importing a text-product
// module. Nothing here knows any document model.

/** The three search options, each persisted for the session by the view. */
export interface FindOptions {
  matchCase: boolean;
  wholeWords: boolean;
  regex: boolean;
}

/** A successfully compiled query — a global RegExp findMatches iterates. */
export interface CompiledQuery {
  regExp: RegExp;
}

/** A query that could not be compiled (an invalid regular expression), so the
 * search must not run and the bar shows the message inline (AC). */
export interface QueryError {
  error: string;
}

export const INVALID_REGEX_MESSAGE = "Invalid regular expression";

/** Narrows a compileQuery result to the error case (a `.vue` template can't
 * write the `in` narrowing itself; exported for symmetry with the view). */
export function isQueryError(value: CompiledQuery | QueryError): value is QueryError {
  return (value as QueryError).error !== undefined;
}

/** Escapes a literal string so it matches verbatim when used as a regex
 * source (non-regex search). */
export function escapeRegExp(literal: string): string {
  return literal.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Compiles a query string under the current options into a global RegExp, or
 * an error for an invalid regular expression (AC: invalid regex shows an
 * inline error and does not run). Returns null for an empty query — there is
 * nothing to search for, which the view renders as "0 of 0" rather than an
 * error. A non-regex query is escaped, so only regex mode can fail to
 * compile. "Whole words only" wraps the source in word boundaries, combining
 * with every mode. */
export function compileQuery(
  query: string,
  options: FindOptions,
): CompiledQuery | QueryError | null {
  if (query === "") {
    return null;
  }
  let source = options.regex ? query : escapeRegExp(query);
  if (options.wholeWords) {
    source = `\\b(?:${source})\\b`;
  }
  const flags = options.matchCase ? "g" : "gi";
  try {
    return { regExp: new RegExp(source, flags) };
  } catch {
    return { error: INVALID_REGEX_MESSAGE };
  }
}
