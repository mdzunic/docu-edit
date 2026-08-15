import { beforeEach, describe, expect, it } from 'vitest'
import {
  ALL_CHARACTERS,
  CHARACTER_CATEGORIES,
  GLYPHLESS_CODE_POINTS,
  RECENT_CHARACTERS_LIMIT,
  RECENT_CHARACTERS_STORAGE_KEY,
  displayName,
  findCharacter,
  isGlyphless,
  loadRecentCharacters,
  matchesQuery,
  parseHexQuery,
  pushRecentCharacter,
  searchCharacters,
  type CategoryId,
} from './specialCharacters'

const EXPECTED_COUNTS: Record<CategoryId, number> = {
  punctuation: 48,
  currency: 38,
  maths: 64,
  arrows: 32,
  latin: 190,
  greek: 52,
  symbols: 64,
  emoji: 64,
}

describe('character catalogue', () => {
  it('has exactly the eight categories', () => {
    expect(CHARACTER_CATEGORIES.map((c) => c.id)).toEqual([
      'punctuation',
      'currency',
      'maths',
      'arrows',
      'latin',
      'greek',
      'symbols',
      'emoji',
    ])
  })

  it('has the per-category counts from the spec', () => {
    for (const category of CHARACTER_CATEGORIES) {
      expect(category.chars.length).toBe(EXPECTED_COUNTS[category.id])
    }
    expect(ALL_CHARACTERS.length).toBe(552)
  })

  it('has a unique code point for every character across the catalogue', () => {
    const seen = new Set<number>()
    for (const char of ALL_CHARACTERS) {
      expect(seen.has(char.cp)).toBe(false)
      seen.add(char.cp)
    }
    expect(seen.size).toBe(552)
  })

  it('has a non-empty name for every character', () => {
    for (const char of ALL_CHARACTERS) {
      expect(char.name.length).toBeGreaterThan(0)
    }
  })

  it('orders every non-emoji category by code point', () => {
    for (const category of CHARACTER_CATEGORIES) {
      if (category.id === 'emoji') continue
      const cps = category.chars.map((c) => c.cp)
      const sorted = [...cps].sort((a, b) => a - b)
      expect(cps).toEqual(sorted)
    }
  })

  it('keeps the thematic emoji order (faces first, food last)', () => {
    const emoji = CHARACTER_CATEGORIES.find((c) => c.id === 'emoji')!
    expect(emoji.chars[0].cp).toBe(0x1f600) // GRINNING FACE
    // Not in code-point order (a later entry has a lower code point).
    const cps = emoji.chars.map((c) => c.cp)
    const sorted = [...cps].sort((a, b) => a - b)
    expect(cps).not.toEqual(sorted)
  })

  it('contains a few known code points and excludes the ones moved elsewhere', () => {
    expect(findCharacter(0x2014)?.name).toBe('EM DASH')
    expect(findCharacter(0x00d7)).toBeDefined() // multiplication sign -> Maths
    // Latin excludes 00D7 / 00F7 (they belong to Maths).
    const latin = CHARACTER_CATEGORIES.find((c) => c.id === 'latin')!
    expect(latin.chars.some((c) => c.cp === 0x00d7)).toBe(false)
    expect(latin.chars.some((c) => c.cp === 0x00f7)).toBe(false)
    // Greek skips the unassigned 03A2.
    const greek = CHARACTER_CATEGORIES.find((c) => c.id === 'greek')!
    expect(greek.chars.some((c) => c.cp === 0x03a2)).toBe(false)
  })

  it('marks glyphless code points', () => {
    for (const cp of [0x00a0, 0x00ad, 0x200b, 0x2002, 0x2003, 0x2009, 0x202f]) {
      expect(isGlyphless(cp)).toBe(true)
      expect(GLYPHLESS_CODE_POINTS.has(cp)).toBe(true)
    }
    expect(isGlyphless(0x2014)).toBe(false)
  })
})

describe('search', () => {
  it('filters by name with case-insensitive substring matching', () => {
    const hits = searchCharacters('em dash')
    expect(hits.some((c) => c.cp === 0x2014)).toBe(true)
    expect(searchCharacters('EM DASH')).toEqual(hits)
  })

  it('filters by keyword alias', () => {
    expect(searchCharacters('nbsp').some((c) => c.cp === 0x00a0)).toBe(true)
    expect(searchCharacters('copyright').some((c) => c.cp === 0x00a9)).toBe(true)
    expect(searchCharacters('tick').some((c) => c.cp === 0x2713)).toBe(true)
  })

  it('filters by hex code point in every accepted form', () => {
    for (const q of ['2014', 'u+2014', 'U+2014', '0x2014', '0X2014', '#2014', '02014']) {
      const hits = searchCharacters(q)
      expect(hits.length).toBe(1)
      expect(hits[0].cp).toBe(0x2014)
    }
  })

  it('matches code points exactly, not as a substring', () => {
    // "201" is a hex prefix of many code points but is not a bare match for
    // 0x2014; it should not hex-match it (0x201 is not in the catalogue).
    expect(searchCharacters('201').some((c) => c.cp === 0x2014)).toBe(false)
  })

  it('parseHexQuery accepts the documented forms and rejects names', () => {
    expect(parseHexQuery('2014')).toBe(0x2014)
    expect(parseHexQuery('u+2014')).toBe(0x2014)
    expect(parseHexQuery('0x2014')).toBe(0x2014)
    expect(parseHexQuery('#2014')).toBe(0x2014)
    expect(parseHexQuery('00a0')).toBe(0x00a0)
    expect(parseHexQuery('dash')).toBeNull()
    expect(parseHexQuery('em dash')).toBeNull()
  })

  it('returns every character for an empty query', () => {
    expect(searchCharacters('').length).toBe(ALL_CHARACTERS.length)
    expect(matchesQuery(ALL_CHARACTERS[0], '')).toBe(true)
  })
})

describe('displayName', () => {
  it('title-cases a Unicode name', () => {
    expect(displayName('EM DASH')).toBe('Em Dash')
    expect(displayName('NO-BREAK SPACE')).toBe('No-Break Space')
  })
})

describe('recent characters', () => {
  beforeEach(() => {
    localStorage.clear()
  })

  it('starts empty and persists pushed characters', () => {
    expect(loadRecentCharacters()).toEqual([])
    pushRecentCharacter(0x2014, [])
    expect(loadRecentCharacters()).toEqual([0x2014])
    expect(
      JSON.parse(localStorage.getItem(RECENT_CHARACTERS_STORAGE_KEY)!),
    ).toEqual([0x2014])
  })

  it('is most-recent-first and de-duplicated', () => {
    let recent: number[] = []
    recent = pushRecentCharacter(0x2014, recent)
    recent = pushRecentCharacter(0x2013, recent)
    recent = pushRecentCharacter(0x2014, recent)
    expect(recent).toEqual([0x2014, 0x2013])
  })

  it('caps at RECENT_CHARACTERS_LIMIT', () => {
    let recent: number[] = []
    for (const char of ALL_CHARACTERS.slice(0, RECENT_CHARACTERS_LIMIT + 10)) {
      recent = pushRecentCharacter(char.cp, recent)
    }
    expect(recent.length).toBe(RECENT_CHARACTERS_LIMIT)
  })

  it('drops code points no longer in the catalogue on load', () => {
    localStorage.setItem(RECENT_CHARACTERS_STORAGE_KEY, JSON.stringify([0x2014, 0x41, 'x']))
    expect(loadRecentCharacters()).toEqual([0x2014])
  })
})
