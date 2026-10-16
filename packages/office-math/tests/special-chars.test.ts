import { describe, expect, it } from 'vitest'
import { SPECIAL_CHARACTERS, autoCorrectReplacement, symbolSearchText } from '../src/index'

describe('special characters', () => {
  it('lists unique characters whose names match Unicode', () => {
    expect(new Set(SPECIAL_CHARACTERS.map((c) => c.id)).size).toBe(SPECIAL_CHARACTERS.length)
    for (const item of SPECIAL_CHARACTERS) {
      expect([...item.char]).toHaveLength(1)
      expect(symbolSearchText(item.char)).toContain('U+')
    }
  })
})

describe('autoCorrectReplacement', () => {
  it('applies defaults after whitespace or at the line start only', () => {
    expect(autoCorrectReplacement('(c)')).toEqual({ length: 3, with: '©' })
    expect(autoCorrectReplacement('Fee (tm)')).toEqual({ length: 4, with: '™' })
    expect(autoCorrectReplacement('abc(c)')).toBeNull()
    expect(autoCorrectReplacement('wait...')).toBeNull()
    expect(autoCorrectReplacement('wait ...')).toEqual({ length: 3, with: '…' })
  })
  it('lets user entries override defaults and prefers the longest match', () => {
    expect(autoCorrectReplacement('x (c)', { '(c)': 'C' })).toEqual({ length: 3, with: 'C' })
    expect(autoCorrectReplacement('x -->', { '>': '!' })).toEqual({ length: 3, with: '→' })
  })
})
