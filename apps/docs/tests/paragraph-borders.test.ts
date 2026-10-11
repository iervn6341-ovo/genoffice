import { describe, expect, it } from 'vitest'
import {
  bordersOf,
  hasSide,
  parseBorderLines,
  presetBorders,
  toggleSide,
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
