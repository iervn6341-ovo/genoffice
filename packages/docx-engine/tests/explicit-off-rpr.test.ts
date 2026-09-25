import { describe, expect, it } from 'vitest'
import { mergeRPrModel } from '../src/generate'

// C1: explicit w:b/w:i val=0 is the only way to switch off style-inherited bold/italic
describe('mergeRPrModel explicit off (C1)', () => {
  it('writes the off-switch over raw rPr that had none', () => {
    const out = mergeRPrModel(
      '<w:rPr><w:color w:val="FF0000"/></w:rPr>',
      { text: 'x', bold: false, color: 'FF0000' },
      false,
    )
    expect(out).toContain('<w:b w:val="0"/>')
    expect(out).toContain('<w:color w:val="FF0000"/>')
  })
  it('keeps a raw off-switch byte-identical for an unset or false model', () => {
    const raw = '<w:rPr><w:b w:val="0"/></w:rPr>'
    expect(mergeRPrModel(raw, { text: 'x' }, false)).toBe(raw)
    expect(mergeRPrModel(raw, { text: 'x', bold: false }, false)).toBe(raw)
  })
  it('replaces raw bold with an explicit off', () => {
    const out = mergeRPrModel('<w:rPr><w:b/></w:rPr>', { text: 'x', bold: false }, false)
    expect(out).toContain('<w:b w:val="0"/>')
    expect(out).not.toContain('<w:b/>')
  })
})
