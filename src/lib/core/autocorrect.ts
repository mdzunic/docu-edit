// Smart-quote and dash autocorrect rules for SPEC-026. Pure functions over
// (prev, typed character, settings) — no Vue imports, no model access. The
// caller (DocumentView) owns the caret/model and the small amount of tracking
// state these decisions read (which substitution was produced at which caret
// position); this module only decides *what* the substitution is.

export interface AutocorrectSettings {
  smartQuotes: boolean
  dashes: boolean
}

export const AUTOCORRECT_STORAGE_KEY = 'docu-edit:autocorrect'

/** Both substitutions default to on (SPEC-026). */
export const DEFAULT_AUTOCORRECT_SETTINGS: AutocorrectSettings = {
  smartQuotes: true,
  dashes: true,
}

/** Load persisted settings, tolerating absent/corrupt storage (private mode,
 * quota, hand-edited JSON) by falling back to the defaults per field. */
export function loadAutocorrectSettings(): AutocorrectSettings {
  try {
    const raw = localStorage.getItem(AUTOCORRECT_STORAGE_KEY)
    if (!raw) return { ...DEFAULT_AUTOCORRECT_SETTINGS }
    const parsed: unknown = JSON.parse(raw)
    if (typeof parsed !== 'object' || parsed === null) {
      return { ...DEFAULT_AUTOCORRECT_SETTINGS }
    }
    const obj = parsed as Record<string, unknown>
    return {
      smartQuotes:
        typeof obj.smartQuotes === 'boolean'
          ? obj.smartQuotes
          : DEFAULT_AUTOCORRECT_SETTINGS.smartQuotes,
      dashes:
        typeof obj.dashes === 'boolean'
          ? obj.dashes
          : DEFAULT_AUTOCORRECT_SETTINGS.dashes,
    }
  } catch {
    return { ...DEFAULT_AUTOCORRECT_SETTINGS }
  }
}

export function saveAutocorrectSettings(settings: AutocorrectSettings): void {
  try {
    localStorage.setItem(AUTOCORRECT_STORAGE_KEY, JSON.stringify(settings))
  } catch {
    // Persistence is best-effort; the in-memory setting still applies.
  }
}

// --- smart quotes ----------------------------------------------------------

export const LEFT_SINGLE_QUOTE = '‘'
export const RIGHT_SINGLE_QUOTE = '’'
export const LEFT_DOUBLE_QUOTE = '“'
export const RIGHT_DOUBLE_QUOTE = '”'
export const EN_DASH = '–'
export const EM_DASH = '—'

/**
 * The code points that make a typed quote an *opening* quote when they sit
 * immediately before it: whitespace, the LineBreak inline (a literal "\n"),
 * the opening brackets/punctuation `( [ { < ¿ ¡ « ‹ / \ – — ‑ …`, and the
 * opening curly quotes. Everything else is a closing context.
 */
export const OPENING_BEFORE: ReadonlySet<number> = new Set<number>([
  // whitespace
  0x0020, 0x0009, 0x00a0, 0x2000, 0x2001, 0x2002, 0x2003, 0x2004, 0x2005,
  0x2006, 0x2007, 0x2008, 0x2009, 0x200a, 0x202f, 0x205f, 0x3000,
  // a forced line break is stored as U+000A
  0x000a,
  // opening brackets / punctuation
  0x0028, 0x005b, 0x007b, 0x003c, 0x00bf, 0x00a1, 0x00ab, 0x2039, 0x002f,
  0x005c, 0x2013, 0x2014, 0x2011, 0x2026,
  // opening curly quotes
  0x2018, 0x201c,
])

/**
 * True when `prev` (the code point immediately before the insertion point, or
 * null at a paragraph start) is an opening context for a quote.
 */
export function isOpeningBefore(prev: string | null): boolean {
  if (prev === null || prev === '') return true
  const cp = prev.codePointAt(0)
  return cp !== undefined && OPENING_BEFORE.has(cp)
}

/**
 * The typographic quote a typed straight quote becomes. Returns null when
 * `typed` is not a straight quote. The choice depends only on `prev`.
 */
export function smartQuote(typed: string, prev: string | null): string | null {
  if (typed === "'") {
    return isOpeningBefore(prev) ? LEFT_SINGLE_QUOTE : RIGHT_SINGLE_QUOTE
  }
  if (typed === '"') {
    return isOpeningBefore(prev) ? LEFT_DOUBLE_QUOTE : RIGHT_DOUBLE_QUOTE
  }
  return null
}

// --- dashes ----------------------------------------------------------------

/** A range in a paragraph's plain text to replace with `text`. */
export interface Replacement {
  start: number
  end: number
  text: string
}

/**
 * Decide the dash substitution for a `-` just typed at `caret` in `text`
 * (both already reflecting the plain insertion of the `-`). `producedEnDashAt`
 * is the caret offset immediately after an en dash this feature produced and
 * has not since been invalidated, or null.
 *
 * - Third `-` after our en dash: replace the en dash and the typed `-` with an
 *   em dash (U+2014).
 * - Otherwise, a `--` whose preceding character is not `-`: replace both with
 *   an en dash (U+2013).
 */
export function dashSubstitution(
  text: string,
  caret: number,
  producedEnDashAt: number | null,
): { replacement: Replacement; producesEnDash: boolean } | null {
  if (caret < 2) return null
  // Trigger 2 — em dash from our en dash plus the typed hyphen.
  if (
    producedEnDashAt === caret - 1 &&
    text[caret - 2] === EN_DASH &&
    text[caret - 1] === '-'
  ) {
    return {
      replacement: { start: caret - 2, end: caret, text: EM_DASH },
      producesEnDash: false,
    }
  }
  // Trigger 1 — en dash from `--`, unless the char before the pair is `-`.
  if (
    text[caret - 2] === '-' &&
    text[caret - 1] === '-' &&
    (caret - 3 < 0 || text[caret - 3] !== '-')
  ) {
    return {
      replacement: { start: caret - 2, end: caret, text: EN_DASH },
      producesEnDash: true,
    }
  }
  return null
}

export function isDigit(ch: string): boolean {
  return ch.length === 1 && ch >= '0' && ch <= '9'
}

/**
 * The decade-idiom correction: a digit typed immediately after a U+2018 this
 * feature just produced flips that opening single quote to U+2019 (so `'90s`
 * gives `’90s`). Returns the replacement of just the quote, or null. Unlike
 * the quote/dash substitutions this is part of the digit's own edit.
 */
export function decadeSubstitution(
  text: string,
  caret: number,
  typed: string,
  producedOpenSingleAt: number | null,
): Replacement | null {
  if (!isDigit(typed)) return null
  if (producedOpenSingleAt !== caret - 1) return null
  if (text[caret - 2] !== LEFT_SINGLE_QUOTE) return null
  return { start: caret - 2, end: caret - 1, text: RIGHT_SINGLE_QUOTE }
}
