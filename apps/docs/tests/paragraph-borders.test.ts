import { describe, expect, it } from 'vitest'
import {
  bordersOf,
  hasSide,
  linesWithWordSpacing,
  parseBorderLines,
  presetBorders,
  toggleSide,
  WORD_BORDER_SPACE_PT,
} from '../src/renderer/editor/paragraph-borders'
import { paraBorderCss } from '../src/renderer/editor/hf-dom'

describe('paragraph border sides (Word Home → Borders)', () => {
  it('each side toggles on its own and the entries combine', () => {
    let b: string | null = null
    b = toggleSide(b, 'b')
    expect(b).toBe('b')
    b = toggleSide(b, 't')
    expect(b).toBe('tb')
    expect(hasSide(b, 't')).toBe(true)
    b = toggleSide(b, 'b')
    expect(b).toBe('t')
    expect(toggleSide(b, 't')).toBeNull()
  })

  it('sides always come out in the model order t b l r h', () => {
    expect(bordersOf(['r', 'h', 'b', 't', 'l'])).toBe('tblrh')
    expect(bordersOf([])).toBeNull()
  })

  it('presets: all = box + between, outside = box only, inside adds just the between rule', () => {
    expect(presetBorders(null, 'all')).toBe('tblrh')
    expect(presetBorders('tblrh', 'outside')).toBe('tblr')
    expect(presetBorders('b', 'inside')).toBe('bh')
    expect(presetBorders('tblr', 'none')).toBeNull()
  })

  it('reads the per-side line JSON and ignores malformed input', () => {
    expect(parseBorderLines('{"b":{"color":"FF0000","szPt":1,"val":"double"}}')).toEqual({
      b: { color: 'FF0000', szPt: 1, val: 'double' },
    })
    expect(parseBorderLines('not json')).toEqual({})
    expect(parseBorderLines(null)).toEqual({})
  })

  it('draws the declared line style as CSS', () => {
    expect(paraBorderCss({ color: 'FF0000', szPt: 1 })).toBe('1px solid #FF0000')
    expect(paraBorderCss({ color: 'FF0000', szPt: 1, val: 'dotted' })).toContain('dotted #FF0000')
    expect(paraBorderCss({ szPt: 1.5, val: 'dashed' })).toMatch(/px dashed #444$/)
    expect(paraBorderCss({ szPt: 3, val: 'double' })).toMatch(/px double #444$/)
  })
})

describe("Word's gap between a paragraph border and its text (Home → Borders, Word 16)", () => {
  it('is 1 pt above, below and between, and 4 pt beside', () => {
    expect(WORD_BORDER_SPACE_PT).toEqual({ t: 1, b: 1, h: 1, l: 4, r: 4 })
  })

  it('gives each side a border adds that gap, as a saved Word document has it', () => {
    // Word saved: bottom → space=1; box → top 1, left 4, bottom 1, right 4; box + between → between 1
    expect(JSON.parse(linesWithWordSpacing('b', null)!)).toEqual({ b: { spacePt: 1 } })
    expect(JSON.parse(linesWithWordSpacing('tblr', null)!)).toEqual({
      t: { spacePt: 1 },
      b: { spacePt: 1 },
      l: { spacePt: 4 },
      r: { spacePt: 4 },
    })
    expect(JSON.parse(linesWithWordSpacing('tblrh', null)!).h).toEqual({ spacePt: 1 })
  })

  it('keeps the line a side already has (its weight, color and its own gap)', () => {
    const had = JSON.stringify({ b: { szPt: 2, color: 'FF0000', spacePt: 6 } })
    expect(JSON.parse(linesWithWordSpacing('bt', had)!)).toEqual({
      b: { szPt: 2, color: 'FF0000', spacePt: 6 },
      t: { spacePt: 1 },
    })
  })

  it('is null when no border is drawn', () => {
    expect(linesWithWordSpacing(null, JSON.stringify({ b: { spacePt: 1 } }))).toBeNull()
    expect(linesWithWordSpacing('', null)).toBeNull()
  })
})
