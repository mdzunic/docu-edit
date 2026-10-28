// Where a library search result opens (SPEC-114): AT its match, in the
// document's own find surface. The app layer, where the products compose:
// a text document opens at /doc/:id?find=<text> (SPEC-017's find bar), a
// deck at /slides/:id?find=<text> (SPEC-098's deck find), a PDF at
// /media/:id?find=<text> (shown above the browser's viewer). When the
// result's first match is not the first one that find surface would select,
// `n` names which one it is (1-based), so the surface selects the nth match —
// counted by the surface's own matcher over the stored model, from the body
// offset the index reported. Loaded on demand by HomeView when a result is
// opened, so none of this rides the library's initial bundle.
import type { RouteLocationRaw } from "vue-router";
import { documentKindOf, getDocument, getPresentation } from "./documentStore";
import { editorPathFor } from "./documentKinds";
import { documentText, documentTextCaret } from "./documentText";
import { compileQuery, findMatches, indexAtOrAfter, isQueryError } from "./findReplace";
import { deckText, deckTextSlide } from "./presentation/deckText";
import { findInDeck } from "./presentation/deckSearch";

/** The find options a deep link compiles its text under: the find bar's own
 * defaults (case-insensitive, substrings, literal). */
const LINK_FIND_OPTIONS = { matchCase: false, wholeWords: false, regex: false } as const;

/** Which match (1-based) of `find` the result's first body match is, as the
 * document's own find surface counts them; 1 when it cannot tell. */
export function matchOrdinal(id: string, find: string, at: number): number {
  const compiled = compileQuery(find, LINK_FIND_OPTIONS);
  if (compiled === null || isQueryError(compiled) || at < 0) {
    return 1;
  }
  const kind = documentKindOf(id);
  if (kind === "text") {
    const model = getDocument(id);
    const caret = model ? documentTextCaret(model, documentText(model), at) : null;
    if (!model || !caret) {
      return 1;
    }
    return Math.max(0, indexAtOrAfter(findMatches(model, compiled), caret)) + 1;
  }
  if (kind === "presentation") {
    const model = getPresentation(id);
    const slide = model ? deckTextSlide(deckText(model), at) : -1;
    if (!model || slide < 0) {
      return 1;
    }
    const index = findInDeck(model, compiled).findIndex((hit) => hit.slideIndex >= slide);
    return index < 0 ? 1 : index + 1;
  }
  return 1;
}

/** The route a search result opens: its editor's path, with `find` (and `n`
 * past the first match) when the search has words to find. */
export function searchResultLocation(id: string, find: string, at: number): RouteLocationRaw {
  const path = editorPathFor(documentKindOf(id) ?? "text", id);
  const text = find.trim();
  if (text === "") {
    return path;
  }
  const n = matchOrdinal(id, text, at);
  return { path, query: n > 1 ? { find: text, n: String(n) } : { find: text } };
}
