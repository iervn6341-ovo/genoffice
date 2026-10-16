import { describe, expect, it } from 'vitest'
import { SPECIAL_CHARS, specialCharGlyph } from '../src/renderer/special-chars'

describe('special characters', () => {
  it('lists unique single code point characters with names', () => {
    const chars = SPECIAL_CHARS.map((s) => s.char)
    expect(new Set(chars).size).toBe(chars.length)
    for (const s of SPECIAL_CHARS) {
      expect([...s.char]).toHaveLength(1)
      expect(s.name).toMatch(/^[A-Z0-9 -]+$/)
    }
  })

  it('shows invisible characters with a visible stand-in', () => {
    expect(specialCharGlyph('\u00a0')).not.toBe('\u00a0')
    expect(specialCharGlyph('\u2014')).toBe('\u2014')
  })
})
