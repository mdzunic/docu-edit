// Word count and document statistics (SPEC-036). Framework-free like the
// rest of src/lib: a tokenizer over plain paragraph text plus a walk of the
// document's block tree, with a per-block cache keyed by block object
// identity so a keystroke re-counts only the paragraph it touched — the
// same "measure once, cache by identity, replaced blocks miss" contract
// lib/pagination's measurement cache relies on (see usePagination.ts),
// reused here instead of inventing a second caching scheme.
//
// The walk only ever visits `Section.blocks` (the body) and, when asked,
// the document's *referenced* footnotes/endnotes (lib/notes.ts's
// referencedNotes) — never `Section.headers`/`Section.footers`, which live
// in separate maps the walk never reaches, so header/footer text is
// excluded by construction rather than by a filter that could drift out of
// sync. Comments (SPEC-038) attach via a side-list the walk never reaches,
// so there is nothing extra to exclude for those. Tracked deletions
// (SPEC-039) are ordinary runs still sitting in `Paragraph.runs` (retained,
// struck through, not removed) — visibleBlockText (lib/trackChanges.ts) is
// blockText with those runs skipped, which is what this module counts with
// instead (AC "Tracked deletions are excluded from word count").
import {
  isBreak,
  isImage,
  isTable,
  isToc,
  isTof,
  type Block,
  type DocumentModel,
} from "./documentModel";
import { blockText, paragraphAt, paragraphSlicesInRange } from "./textEditing";
import { isCollapsed, orderedRange, type ModelSelection } from "./commands";
import { referencedNotes } from "./notes";
import { visibleBlockText } from "./trackChanges";

/** Character-classification counts derived straight from a text sample —
 * everything countBlock/countText can compute without knowing whether the
 * text belongs to a whole paragraph (paragraphs is block-level, added by
 * the caller). */
export interface TextCounts {
  words: number;
  charactersWithSpaces: number;
  charactersWithoutSpaces: number;
  sentences: number;
  /** CJK (Han ideograph / Hiragana / Katakana / Hangul) characters, counted
   * individually rather than grouped into a Latin-style word (AC "CJK
   * characters are counted as characters with a separate CJK count
   * shown"). Included in `words` too — each is also its own word, matching
   * word processors' Asian-character convention. */
  cjkCharacters: number;
}

/** Block/document-level statistics: TextCounts plus a paragraph tally. */
export interface CountStats extends TextCounts {
  paragraphs: number;
}

export function emptyCountStats(): CountStats {
  return {
    words: 0,
    charactersWithSpaces: 0,
    charactersWithoutSpaces: 0,
    sentences: 0,
    cjkCharacters: 0,
    paragraphs: 0,
  };
}

export function addCountStats(a: CountStats, b: CountStats): CountStats {
  return {
    words: a.words + b.words,
    charactersWithSpaces: a.charactersWithSpaces + b.charactersWithSpaces,
    charactersWithoutSpaces: a.charactersWithoutSpaces + b.charactersWithoutSpaces,
    sentences: a.sentences + b.sentences,
    cjkCharacters: a.cjkCharacters + b.cjkCharacters,
    paragraphs: a.paragraphs + b.paragraphs,
  };
}

// --- tokenizing ---------------------------------------------------------

/** The scripts word processors group under "Asian character count": Han
 * ideographs (plus the Extension A and Compatibility blocks), Hiragana,
 * Katakana and Hangul syllables. Each character in these ranges is its own
 * countable unit rather than part of a Latin-style word. */
const CJK_CHAR =
  /[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\u3040-\u30FF\uAC00-\uD7A3]/u;

/** Punctuation a sentence hangs on a token's end ("see example.com.") —
 * peeled off a URL token before it is counted, matching the trimming
 * lib/links.ts's detectBareUrl already applies when autolinking. */
const TRAILING_PUNCTUATION = /[.,;:!?)\]}>"']+$/;

/** Recognises an http(s)/www. URL token during tokenizing. */
const URL_PREFIX = /^(?:https?:\/\/|www\.)/i;

/** One tokenizing pass over a text sample, in document order: an http(s)/
 * www. URL (AC "URLs count as one word"), a single CJK character (AC "CJK
 * characters are counted as characters"), or a maximal run of letters and
 * numbers that may contain — but never start or end with — an internal
 * hyphen or apostrophe (AC "hyphenated words count as one word" and
 * "numbers count as words"): "well-known" and "42" are each one word,
 * while a hyphen used as a dash between separate words ("word - word")
 * does not merge them because the space breaks the run. */
const TOKEN_PATTERN =
  /(?:https?:\/\/|www\.)\S+|[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\u3040-\u30FF\uAC00-\uD7A3]|[\p{L}\p{N}](?:[\p{L}\p{N}'\u2019-]*[\p{L}\p{N}])?/gu;

/** Counts words, characters and sentences in a plain text sample (a
 * paragraph's concatenated run text, or a selection's). Character counts
 * are UTF-16 code units, matching how the rest of the model already
 * measures text (CaretPosition.offset, Run.text.length — see
 * textEditing.blockLength) rather than grapheme clusters. */
export function countText(text: string): TextCounts {
  let words = 0;
  let cjkCharacters = 0;
  for (const match of text.matchAll(TOKEN_PATTERN)) {
    const token = match[0];
    if (CJK_CHAR.test(token) && token.length === 1) {
      words += 1;
      cjkCharacters += 1;
      continue;
    }
    if (URL_PREFIX.test(token)) {
      if (token.replace(TRAILING_PUNCTUATION, "").length > 0) {
        words += 1;
      }
      continue;
    }
    words += 1;
  }
  return {
    words,
    charactersWithSpaces: text.length,
    charactersWithoutSpaces: text.replace(/\s/g, "").length,
    sentences: countSentences(text),
    cjkCharacters,
  };
}

/** A sentence terminator: one or more `.!?` immediately before whitespace
 * or the end of the text — so "3.14" and "e.g. " (period followed by a
 * letter, not by a space) don't end a sentence, but a run of punctuation
 * ("Really?!" or an ellipsis) ends exactly one. A heuristic, not a full
 * abbreviation-aware sentence splitter — SPEC-036 asks for a count shown
 * in the dialog, not linguistic precision. */
const SENTENCE_TERMINATOR = /[.!?]+(?=\s|$)/g;

function countSentences(text: string): number {
  const trimmed = text.trim();
  if (trimmed.length === 0) {
    return 0;
  }
  const matches = [...trimmed.matchAll(SENTENCE_TERMINATOR)];
  let count = matches.length;
  const last = matches[matches.length - 1];
  const tailStart = last ? last.index + last[0].length : 0;
  if (trimmed.slice(tailStart).trim().length > 0) {
    // Trailing text with no closing punctuation still counts as a sentence.
    count += 1;
  }
  return count;
}

// --- block/document walk -------------------------------------------------

/** One block's statistics: for a table, the summed statistics of every
 * cell paragraph (the same recursive walk blockText uses); for an image,
 * break or table of contents, nothing (they carry no countable text); for
 * a paragraph, its visible text's counts (tracked deletions excluded,
 * SPEC-039 AC-17) plus one paragraph. */
export function countBlock(block: Block): CountStats {
  if (isTable(block)) {
    let stats = emptyCountStats();
    for (const row of block.rows) {
      for (const cell of row.cells) {
        for (const paragraph of cell.blocks) {
          stats = addCountStats(stats, countBlock(paragraph));
        }
      }
    }
    return stats;
  }
  if (isImage(block) || isBreak(block) || isToc(block) || isTof(block)) {
    return emptyCountStats();
  }
  return { ...countText(visibleBlockText(block)), paragraphs: 1 };
}

/** A block's cached statistics, computing and storing on a miss. `cache`
 * is keyed by block object identity: commands.ts's mutation contract
 * replaces an edited block rather than mutating it in place, so an
 * unchanged block always hits and a changed one always misses — the exact
 * invariant lib/pagination's measurement cache already relies on. */
function cachedBlockCounts(block: Block, cache: WeakMap<Block, CountStats>): CountStats {
  let counts = cache.get(block);
  if (!counts) {
    counts = countBlock(block);
    cache.set(block, counts);
  }
  return counts;
}

/** The document's body statistics: every section's blocks, walked in
 * order. Headers, footers (SPEC-029) and notes (SPEC-031) are never
 * reached from here — see the module header. */
export function documentBodyStats(
  model: DocumentModel,
  cache: WeakMap<Block, CountStats> = new WeakMap(),
): CountStats {
  let stats = emptyCountStats();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      stats = addCountStats(stats, cachedBlockCounts(block, cache));
    }
  }
  return stats;
}

/** The document's referenced footnote/endnote statistics (SPEC-031) — the
 * "Include footnotes and endnotes" toggle's contribution. Uses
 * lib/notes.ts's referencedNotes rather than `model.notes` directly: a
 * note whose reference mark was deleted stays in that list but must stay
 * invisible everywhere, counting included (see NoteContent's own doc
 * comment). */
export function documentNotesStats(
  model: DocumentModel,
  cache: WeakMap<Block, CountStats> = new WeakMap(),
): CountStats {
  let stats = emptyCountStats();
  for (const kind of ["footnote", "endnote"] as const) {
    for (const note of referencedNotes(model, kind)) {
      for (const block of note.blocks) {
        stats = addCountStats(stats, cachedBlockCounts(block, cache));
      }
    }
  }
  return stats;
}

/** The document's full statistics for the statistics dialog: body always,
 * notes only when `includeNotes` is set (AC "Include footnotes and
 * endnotes" toggle). Body and notes share one cache — a block cannot be
 * both, so there is no risk of one invalidating the other. */
export function documentStats(
  model: DocumentModel,
  options: { includeNotes: boolean },
  cache: WeakMap<Block, CountStats> = new WeakMap(),
): CountStats {
  const stats = documentBodyStats(model, cache);
  return options.includeNotes ? addCountStats(stats, documentNotesStats(model, cache)) : stats;
}

// --- selection-aware count ------------------------------------------------

/** A selection's plain text, paragraphs joined with "\n" — the same shape
 * lib/clipboard's fragmentText produces, computed directly off
 * paragraphSlicesInRange instead of through extractFragment's run-copying
 * (unneeded here: word counting never needs formatting). Empty for a
 * collapsed selection. */
export function selectionText(model: DocumentModel, selection: ModelSelection): string {
  if (isCollapsed(selection)) {
    return "";
  }
  const { start, end } = orderedRange(model, selection);
  const parts: string[] = [];
  for (const { address, from, to } of paragraphSlicesInRange(model, start, end)) {
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
      continue;
    }
    parts.push(blockText(paragraph).slice(from, to));
  }
  return parts.join("\n");
}

/** The word count of a selection (AC "N of M words") — 0 for a collapsed
 * selection. */
export function selectionWordCount(model: DocumentModel, selection: ModelSelection): number {
  const text = selectionText(model, selection);
  return text.length === 0 ? 0 : countText(text).words;
}
