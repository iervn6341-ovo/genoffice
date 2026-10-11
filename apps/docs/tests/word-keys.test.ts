import { describe, expect, it } from 'vitest'

import { wordDeleteRange, wordLeft, wordRight } from '../src/renderer/editor/word-keys'

describe('word units', () => {
  it('moves left to the start of the word before the caret', () => {
    expect(wordLeft('Hello, world!', 12)).toBe(7)
    expect(wordLeft('abc def', 7)).toBe(4)
  })

  it('moves right past the word and its trailing spaces', () => {
    expect(wordRight('abc def', 0)).toBe(4)
  })
})

describe('Option+Backspace range', () => {
  it('takes the space before the word when punctuation follows (Word K02)', () => {
    expect(wordDeleteRange('Hello, world!', 12)).toEqual([6, 12])
  })

  it('keeps the space when the word ends the paragraph or text follows', () => {
    expect(wordDeleteRange('Hello, world', 12)).toEqual([7, 12])
    expect(wordDeleteRange('Hello, world again', 12)).toEqual([7, 12])
  })
})
