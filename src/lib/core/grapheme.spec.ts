import { describe, expect, it } from 'vitest'
import { nextGraphemeBoundary, previousGraphemeBoundary } from './grapheme'

// A non-BMP character (MATHEMATICAL ITALIC SMALL X) is two UTF-16 code units.
const X = '\u{1D465}'
// An emoji ZWJ sequence: family, several code points joined by U+200D.
const FAMILY = '👨‍👩‍👧‍👦'
// A skin-tone modified emoji: base + modifier (two code points, 4 units).
const WAVE = '👋\u{1F3FB}'
// An emoji with a variation selector.
const HEART = '❤️'

describe('previousGraphemeBoundary', () => {
  it('steps over a surrogate pair as one unit', () => {
    const text = `a${X}b` // offsets: a=0, X=1..3, b=3
    expect(previousGraphemeBoundary(text, 3)).toBe(1)
    expect(previousGraphemeBoundary(text, 1)).toBe(0)
  })

  it('never lands inside a surrogate pair', () => {
    const text = X
    expect(previousGraphemeBoundary(text, 2)).toBe(0)
  })

  it('steps over a whole ZWJ family emoji at once', () => {
    const text = `a${FAMILY}b`
    const end = 1 + FAMILY.length
    expect(previousGraphemeBoundary(text, end)).toBe(1)
  })

  it('steps over a skin-tone sequence at once', () => {
    const text = WAVE
    expect(previousGraphemeBoundary(text, WAVE.length)).toBe(0)
  })

  it('steps over a variation-selector sequence at once', () => {
    const text = HEART
    expect(previousGraphemeBoundary(text, HEART.length)).toBe(0)
  })

  it('is a no-op at the start', () => {
    expect(previousGraphemeBoundary('abc', 0)).toBe(0)
  })
})

describe('nextGraphemeBoundary', () => {
  it('steps over a surrogate pair as one unit', () => {
    const text = `a${X}b`
    expect(nextGraphemeBoundary(text, 1)).toBe(3)
    expect(nextGraphemeBoundary(text, 0)).toBe(1)
  })

  it('steps over a whole ZWJ family emoji at once', () => {
    const text = `a${FAMILY}b`
    expect(nextGraphemeBoundary(text, 1)).toBe(1 + FAMILY.length)
  })

  it('steps over a skin-tone sequence at once', () => {
    expect(nextGraphemeBoundary(WAVE, 0)).toBe(WAVE.length)
  })

  it('is a no-op at the end', () => {
    expect(nextGraphemeBoundary('abc', 3)).toBe(3)
  })
})
