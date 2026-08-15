// Grapheme-aware caret stepping for SPEC-026. Plain data, no Vue imports.
//
// Offsets stay UTF-16 code-unit indices into a paragraph's plain text
// (`blockText`) — grapheme awareness changes only the *step size* of caret
// motion and deletion, never the meaning of an offset. A step never lands
// inside a surrogate pair or partway through an emoji ZWJ / variation-selector
// / skin-tone sequence.
//
// Boundaries are computed with `Intl.Segmenter({ granularity: 'grapheme' })`.
// Where `Intl.Segmenter` is unavailable the fallback steps by whole code
// point (never by a lone surrogate).

let cachedSegmenter: Intl.Segmenter | null | undefined

function segmenter(): Intl.Segmenter | null {
  if (cachedSegmenter === undefined) {
    cachedSegmenter =
      typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function'
        ? new Intl.Segmenter(undefined, { granularity: 'grapheme' })
        : null
  }
  return cachedSegmenter
}

/** Is the code unit at `index` the low (trailing) half of a surrogate pair? */
function isLowSurrogateAt(text: string, index: number): boolean {
  const code = text.charCodeAt(index)
  return code >= 0xdc00 && code <= 0xdfff
}

/**
 * The set of grapheme-cluster boundary offsets in `text`, always including 0
 * and `text.length`. Uses `Intl.Segmenter` when available, else code-point
 * boundaries (which still never split a surrogate pair).
 */
function boundaries(text: string): number[] {
  const seg = segmenter()
  const result: number[] = [0]
  if (seg) {
    for (const { index } of seg.segment(text)) {
      if (index > 0) result.push(index)
    }
  } else {
    // Code-point fallback: walk by code point so a surrogate pair (which
    // `for...of` yields as one iteration) counts as a single step.
    for (const ch of text) {
      result.push(result[result.length - 1] + ch.length)
    }
    result.pop() // the loop pushed text.length as its final entry's successor
  }
  if (result[result.length - 1] !== text.length) result.push(text.length)
  return result
}

/**
 * The grapheme boundary at or before `offset` but strictly less than it — the
 * offset a Backspace or Left-arrow from `offset` lands on. Returns `offset`
 * unchanged when it is already 0.
 */
export function previousGraphemeBoundary(text: string, offset: number): number {
  if (offset <= 0) return 0
  if (offset > text.length) return text.length
  const bounds = boundaries(text)
  let prev = 0
  for (const b of bounds) {
    if (b >= offset) break
    prev = b
  }
  // Defensive: never land on the low half of a surrogate pair.
  if (prev > 0 && isLowSurrogateAt(text, prev)) prev -= 1
  return prev
}

/**
 * The grapheme boundary strictly after `offset` — the offset a Delete or
 * Right-arrow from `offset` lands on. Returns `offset` unchanged when it is
 * already at `text.length`.
 */
export function nextGraphemeBoundary(text: string, offset: number): number {
  if (offset >= text.length) return text.length
  if (offset < 0) return 0
  const bounds = boundaries(text)
  for (const b of bounds) {
    if (b > offset) {
      return isLowSurrogateAt(text, b) ? b + 1 : b
    }
  }
  return text.length
}
