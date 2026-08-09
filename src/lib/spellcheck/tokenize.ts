// Splits text into spell-checkable word spans (SPEC-035), skipping URLs and
// email addresses outright (AC-12) rather than tokenizing and filtering them
// after the fact — a URL like "https://a.dev/well-known" must not surface
// "well"/"known" as separate misspelling candidates.

export interface SpellToken {
  text: string;
  /** Offset into the checked text, inclusive. */
  start: number;
  /** Offset into the checked text, exclusive. */
  end: number;
}

// A "word" is one or more letters (any script), allowing an internal
// apostrophe/right-quote or hyphen followed by another letter ("don't",
// "well-known") so contractions and hyphenated compounds are one token, not
// several. \p{Mn} covers combining marks on decomposed accented letters.
const WORD_PATTERN = /\p{L}(?:[\p{L}\p{Mn}]|['’-](?=\p{L}))*/gu;

const URL_SCAN_PATTERN = /\b(?:[a-z][a-z0-9+.-]*:\/\/|www\.)\S+/giu;
const EMAIL_SCAN_PATTERN = /\b[^\s@]+@[^\s@]+\.[a-z]{2,}\b/giu;

function skipSpans(text: string): Array<[number, number]> {
  const spans: Array<[number, number]> = [];
  for (const pattern of [URL_SCAN_PATTERN, EMAIL_SCAN_PATTERN]) {
    pattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text))) {
      spans.push([match.index, match.index + match[0].length]);
    }
  }
  return spans;
}

function withinAnySpan(start: number, end: number, spans: Array<[number, number]>): boolean {
  return spans.some(([spanStart, spanEnd]) => start >= spanStart && end <= spanEnd);
}

/** The spell-checkable word tokens in `text`, in source order, with any URL
 * or email address carved out first (SPEC-035 AC-12). */
export function tokenizeForSpellCheck(text: string): SpellToken[] {
  const spans = skipSpans(text);
  const tokens: SpellToken[] = [];
  WORD_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = WORD_PATTERN.exec(text))) {
    const start = match.index;
    const end = start + match[0].length;
    if (withinAnySpan(start, end, spans)) {
      continue;
    }
    tokens.push({ text: match[0], start, end });
  }
  return tokens;
}
