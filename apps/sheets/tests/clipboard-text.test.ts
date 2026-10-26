import { CellValueType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { isGridText, parseDelimitedText, typeField } from '../src/renderer/clipboard-text'

describe('parseDelimitedText', () => {
  it('splits at tabs and line breaks', () => {
    expect(parseDelimitedText('a\tb\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
    expect(parseDelimitedText('a\tb\r\nc\td\r\n')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
  })

  it('keeps leading, trailing and middle empty cells', () => {
    expect(parseDelimitedText('\ta\t\n\t\tb')).toEqual([
      ['', 'a', ''],
      ['', '', 'b'],
    ])
  })

  it('a final line break does not add a row, a second one adds a blank row', () => {
    expect(parseDelimitedText('a\n')).toEqual([['a']])
    expect(parseDelimitedText('a\n\n')).toEqual([['a'], ['']])
    expect(parseDelimitedText('')).toEqual([['']])
  })

  it('reads quoted fields with tabs, line breaks and doubled quotes', () => {
    expect(parseDelimitedText('"a\tb"\t"line1\nline2"\t"say ""hi"""\nz')).toEqual([
      ['a\tb', 'line1\nline2', 'say "hi"'],
      ['z'],
    ])
    expect(parseDelimitedText('"a\r\nb"\tc')).toEqual([['a\nb', 'c']])
  })

  it('treats a quote that is not a whole field as plain text', () => {
    expect(parseDelimitedText('5" pipe\tx')).toEqual([['5" pipe', 'x']])
    expect(parseDelimitedText('"abc"def\tx')).toEqual([['"abc"def', 'x']])
    expect(parseDelimitedText('"unterminated\tx')).toEqual([['"unterminated', 'x']])
  })

  it('does not split on commas by default (Excel pastes CSV text into one column)', () => {
    expect(parseDelimitedText('a,b,c')).toEqual([['a,b,c']])
    expect(parseDelimitedText('a,b', ',')).toEqual([['a', 'b']])
  })

  it('round-trips the copy side quoting', () => {
    const fields = ['plain', 'tab\there', 'multi\nline', 'quote"inside', '', ' lead']
    const quoted = fields.map((f) => (/[\t\n"]/.test(f) ? `"${f.replace(/"/g, '""')}"` : f))
    expect(parseDelimitedText(quoted.join('\t'))).toEqual([fields])
  })
})

describe('isGridText', () => {
  it('is true only for text with a tab or a line break', () => {
    expect(isGridText('a\tb')).toBe(true)
    expect(isGridText('a\nb')).toBe(true)
    expect(isGridText('a,b')).toBe(false)
    expect(isGridText('plain')).toBe(false)
  })
})

describe('typeField', () => {
  it('types numbers, booleans, formulas and text', () => {
    expect(typeField('12.5')).toMatchObject({ v: 12.5, t: CellValueType.NUMBER })
    expect(typeField('TRUE')).toMatchObject({ v: 1, t: CellValueType.BOOLEAN })
    expect(typeField('false')).toMatchObject({ v: 0, t: CellValueType.BOOLEAN })
    expect(typeField('=A1+B1')).toMatchObject({ f: '=A1+B1', v: null })
    expect(typeField('hello')).toMatchObject({ v: 'hello', t: CellValueType.STRING })
    expect(typeField('')).toMatchObject({ v: null, f: null })
  })

  it('percentages and dates carry their number format', () => {
    expect(typeField('25%')).toMatchObject({ v: 0.25, pattern: expect.stringContaining('%') })
    const date = typeField('2026-01-05')
    expect(typeof date.v).toBe('number')
    expect(date.pattern).toMatch(/y/i)
  })

  it('a leading apostrophe keeps the text, long digit strings stay text', () => {
    expect(typeField("'007")).toMatchObject({ v: '007', t: CellValueType.STRING })
    expect(typeField('12345678901234567890')).toMatchObject({
      v: '12345678901234567890',
      t: CellValueType.STRING,
    })
  })
})
