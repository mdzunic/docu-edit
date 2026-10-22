import { afterEach, describe, expect, it } from 'vitest'
import {
  AUTOCORRECT_STORAGE_KEY,
  DEFAULT_AUTOCORRECT_SETTINGS,
  dashSubstitution,
  decadeSubstitution,
  isDigit,
  isOpeningBefore,
  loadAutocorrectSettings,
  saveAutocorrectSettings,
  smartQuote,
} from './autocorrect'

describe('smart quotes', () => {
  it('opens a single quote at a paragraph start (prev = none)', () => {
    expect(smartQuote("'", null)).toBe('‘')
    expect(smartQuote('"', null)).toBe('“')
  })

  it('opens after whitespace and opening punctuation', () => {
    for (const prev of [' ', ' ', '\n', '(', '[', '«', '/', '–', '…', '‘', '“']) {
      expect(smartQuote("'", prev)).toBe('‘')
      expect(smartQuote('"', prev)).toBe('“')
      expect(isOpeningBefore(prev)).toBe(true)
    }
  })

  it('closes after a letter — the don’t / parents’ rule', () => {
    expect(smartQuote("'", 'n')).toBe('’') // don't
    expect(smartQuote("'", 's')).toBe('’') // parents'
    expect(smartQuote('"', 'd')).toBe('”')
    expect(isOpeningBefore('n')).toBe(false)
  })

  it('closes after a digit and closing bracket', () => {
    expect(smartQuote('"', '5')).toBe('”')
    expect(smartQuote('"', ')')).toBe('”')
  })

  it('returns null for a non-quote character', () => {
    expect(smartQuote('a', null)).toBeNull()
    expect(smartQuote('-', ' ')).toBeNull()
  })
})

describe('dash substitution', () => {
  it('replaces -- with an en dash when the char before is not -', () => {
    // text after typing the 2nd '-': "a--", caret at 3
    const sub = dashSubstitution('a--', 3, null)
    expect(sub).not.toBeNull()
    expect(sub!.replacement).toEqual({ start: 1, end: 3, text: '–' })
    expect(sub!.producesEnDash).toBe(true)
  })

  it('does not fire when the char before the pair is a dash', () => {
    // "a---" caret 4: last two are --, but char before pair is '-'
    expect(dashSubstitution('a---', 4, null)).toBeNull()
  })

  it('promotes our en dash + typed hyphen to an em dash', () => {
    // en dash produced at offset 2 (caret after it), user types '-': "a–-", caret 3
    const sub = dashSubstitution('a–-', 3, 2)
    expect(sub).not.toBeNull()
    expect(sub!.replacement).toEqual({ start: 1, end: 3, text: '—' })
    expect(sub!.producesEnDash).toBe(false)
  })

  it('does not promote when the en dash was not produced at this position', () => {
    expect(dashSubstitution('a–-', 3, null)).toBeNull()
  })

  it('fires en dash at a paragraph start', () => {
    const sub = dashSubstitution('--', 2, null)
    expect(sub!.replacement).toEqual({ start: 0, end: 2, text: '–' })
  })
})

describe('decade idiom', () => {
  it('flips our opening single quote to a closing one before a digit', () => {
    // "‘9" caret 2, our 2018 produced with caret after it at offset 1
    const sub = decadeSubstitution('‘9', 2, '9', 1)
    expect(sub).toEqual({ start: 0, end: 1, text: '’' })
  })

  it('does not fire for a non-digit or a quote we did not produce', () => {
    expect(decadeSubstitution('‘a', 2, 'a', 1)).toBeNull()
    expect(decadeSubstitution('‘9', 2, '9', null)).toBeNull()
  })

  it('isDigit', () => {
    expect(isDigit('0')).toBe(true)
    expect(isDigit('9')).toBe(true)
    expect(isDigit('a')).toBe(false)
    expect(isDigit('12')).toBe(false)
  })
})

describe('settings persistence', () => {
  afterEach(() => localStorage.clear())

  it('defaults both to true', () => {
    expect(DEFAULT_AUTOCORRECT_SETTINGS).toEqual({ smartQuotes: true, dashes: true })
    expect(loadAutocorrectSettings()).toEqual({ smartQuotes: true, dashes: true })
  })

  it('round-trips through localStorage under the documented key', () => {
    saveAutocorrectSettings({ smartQuotes: false, dashes: true })
    expect(JSON.parse(localStorage.getItem(AUTOCORRECT_STORAGE_KEY)!)).toEqual({
      smartQuotes: false,
      dashes: true,
    })
    expect(loadAutocorrectSettings()).toEqual({ smartQuotes: false, dashes: true })
  })

  it('falls back per-field on corrupt data', () => {
    localStorage.setItem(AUTOCORRECT_STORAGE_KEY, '{"smartQuotes":false}')
    expect(loadAutocorrectSettings()).toEqual({ smartQuotes: false, dashes: true })
    localStorage.setItem(AUTOCORRECT_STORAGE_KEY, 'not json')
    expect(loadAutocorrectSettings()).toEqual({ smartQuotes: true, dashes: true })
  })
})
