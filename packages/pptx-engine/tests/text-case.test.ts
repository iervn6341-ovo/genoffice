import { describe, it, expect } from 'vitest'
import { changeTextCase } from '../src/text-case'

const one = (s: string, mode: Parameters<typeof changeTextCase>[1], start = true) =>
  changeTextCase([s], mode, start)[0]

describe('changeTextCase (PowerPoint Change Case)', () => {
  it('lowercase / UPPERCASE', () => {
    expect(one('Hello World', 'lower')).toBe('hello world')
    expect(one('Hello World', 'upper')).toBe('HELLO WORLD')
  })

  it('Sentence case capitalises the start of each sentence and lowercases the rest', () => {
    expect(one('HELLO THERE. how ARE you? fine!  ok', 'sentence')).toBe(
      'Hello there. How are you? Fine!  Ok',
    )
  })

  it('Capitalize Each Word', () => {
    expect(one('the QUICK brown-fox', 'title')).toBe('The Quick Brown-Fox')
  })

  it('tOGGLE cASE swaps every letter', () => {
    expect(one('Hello World', 'toggle')).toBe('hELLO wORLD')
  })

  it('keeps segment boundaries and decides word starts across them', () => {
    // "hel" + "LO wor" + "ld" is one paragraph split over three runs
    expect(changeTextCase(['hel', 'LO wor', 'ld'], 'title')).toEqual(['Hel', 'lo Wor', 'ld'])
    expect(changeTextCase(['a.', ' b'], 'sentence')).toEqual(['A.', ' B'])
  })

  it('a continuation does not start a sentence', () => {
    expect(one('world', 'sentence', false)).toBe('world')
  })

  it('leaves CJK, digits and length-changing mappings untouched', () => {
    expect(one('中文 abc 123', 'upper')).toBe('中文 ABC 123')
    expect(one('straße', 'upper')).toBe('STRAßE')
    expect(one('句子。next', 'sentence')).toBe('句子。Next')
  })
})
