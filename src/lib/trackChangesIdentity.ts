// Author identity and colour for track changes (SPEC-039). The author is
// literally the SPEC-038 settable identity (AC "the author is the SPEC-038
// settable identity") — this module re-exports lib/commentIdentity.ts's
// storage rather than duplicating a second name/preference, so setting your
// name once affects both comments and tracked changes.
export {
  DEFAULT_COMMENT_AUTHOR as DEFAULT_TRACK_CHANGES_AUTHOR,
  loadCommentAuthor as loadTrackChangesAuthor,
  saveCommentAuthor as saveTrackChangesAuthor,
} from "./core/commentIdentity";

// A fixed categorical palette (D3's schemeCategory10) — distinct, readable
// hues for up to ten authors before any repeat. Author -> colour is a pure
// hash, not an encounter-order assignment, so a colour is stable for the
// whole session (and across sessions) regardless of which author's edits
// are opened/scrolled to first (SPEC-039 AC "each author has a stable
// colour for the session").
export const AUTHOR_COLOR_PALETTE: readonly string[] = [
  "#1f77b4",
  "#ff7f0e",
  "#2ca02c",
  "#d62728",
  "#9467bd",
  "#8c564b",
  "#e377c2",
  "#7f7f7f",
  "#bcbd22",
  "#17becf",
];

function hashAuthorName(author: string): number {
  let hash = 0;
  for (let i = 0; i < author.length; i++) {
    hash = (hash * 31 + author.charCodeAt(i)) | 0;
  }
  return Math.abs(hash);
}

/** The colour an author's insertions/deletions/change bars render in
 * (SPEC-039 AC "each author has a stable colour for the session") — a pure
 * function of the author's name, so it never needs session state to stay
 * stable, and two browser tabs/sessions agree on the same author's colour. */
export function authorColor(author: string): string {
  return AUTHOR_COLOR_PALETTE[hashAuthorName(author) % AUTHOR_COLOR_PALETTE.length];
}
