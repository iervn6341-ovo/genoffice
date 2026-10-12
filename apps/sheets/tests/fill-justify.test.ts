import { describe, expect, it } from 'vitest'

import { planJustify, wrapWords } from '../src/renderer/fill-justify'

// a fixed-width font: every character is 7 px wide
const measure = (text: string): number => text.length * 7
const NINE = 9 * 7 + 1 // room for nine characters

describe('Justify breaks the text at spaces into lines no wider than the selection', () => {
  it('fills each line as full as it goes (read from Excel: "The quick" / "brown fox" / "jumps over" …)', () => {
    expect(wrapWords('The quick brown fox jumps over the lazy dog', NINE, measure)).toEqual([
      'The quick',
      'brown fox',
      'jumps',
      'over the',
      'lazy dog',
    ])
  })

  it('puts a word longer than the width on a line of its own', () => {
    expect(wrapWords('a extraordinarily b', NINE, measure)).toEqual(['a', 'extraordinarily', 'b'])
  })

  it('collapses runs of spaces and ignores empty text', () => {
    expect(wrapWords('  a   b  ', 100, measure)).toEqual(['a b'])
    expect(wrapWords('   ', 100, measure)).toEqual([])
  })
})

describe('what Justify does to the selection', () => {
  const rows = (...values: unknown[]) => values.map((v) => [{ v }])

  it('joins the cells of the column with spaces and writes the lines from the first cell down', () => {
    const plan = planJustify(rows('The quick', 'brown fox', 'jumps'), 200, measure)
    expect(plan).toEqual({ lines: ['The quick brown fox jumps'], cleared: 2, extra: 0 })
  })

  it('leaves the cells the lines do not reach empty, and runs under the selection when they need more rows', () => {
    const fewer = planJustify(rows('aa bb', 'cc'), 7 * 20, measure)
    expect(fewer).toMatchObject({ lines: ['aa bb cc'], cleared: 1, extra: 0 })
    const more = planJustify(rows('one two three four five six'), NINE, measure)
    if ('error' in more) throw new Error(more.error)
    expect(more.lines).toEqual(['one two', 'three', 'four five', 'six'])
    expect(more.extra).toBe(3)
    expect(more.cleared).toBe(0)
  })

  it('works on the first column of a wider selection', () => {
    const plan = planJustify(
      [
        [{ v: 'one two' }, { v: '' }, { v: '' }],
        [{ v: 'three' }, {}, {}],
      ],
      7 * 30,
      measure,
    )
    expect(plan).toEqual({ lines: ['one two three'], cleared: 1, extra: 0 })
  })

  it('refuses a selection with a number, a boolean or a formula in it (Excel cannot justify those)', () => {
    expect(planJustify(rows('text', 42, 'more'), 200, measure)).toEqual({ error: 'notText' })
    expect(planJustify(rows('text', true), 200, measure)).toEqual({ error: 'notText' })
    expect(planJustify([[{ v: 'text' }], [{ v: 'x', f: '=A1' }]], 200, measure)).toEqual({
      error: 'notText',
    })
    // a number outside the first column counts too
    expect(planJustify([[{ v: 'text' }, { v: 3 }]], 200, measure)).toEqual({ error: 'notText' })
  })

  it('has nothing to do when no cell holds text', () => {
    expect(planJustify(rows('', null, '   '), 200, measure)).toEqual({ error: 'noText' })
  })
})

describe('Justify leaves formatted text alone', () => {
  it('refuses a selection with a rich-text cell, whose text is not in its plain value', () => {
    const rich = [[{ v: 'plain' }], [{ v: null, p: { id: 'rich', body: { dataStream: 'a\r\n' } } }]]
    expect(planJustify(rich, 200, (text) => text.length * 10)).toEqual({ error: 'rich' })
  })
})
