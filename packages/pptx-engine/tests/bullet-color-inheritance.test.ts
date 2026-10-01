import { describe, expect, it } from 'vitest'
import { parseSlide } from '../src/parse'
import { generateParagraphXml } from '../src/generate'
import type { TextElement } from '../src/types'

function paragraph(defaults: string, local: string) {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    ctx: {},
    slideXml: `<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
      <p:sp><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr>${defaults}</a:lvl1pPr></a:lstStyle>
      <a:p><a:pPr>${local}</a:pPr><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Red first word</a:t></a:r></a:p>
      </p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
  })
  return (slide.elements[0] as TextElement).text!.paragraphs[0]!
}
const brown = '<a:buClr><a:srgbClr val="9E3611"/></a:buClr>'
const glyph = '<a:buChar char="•"/>'

describe('bullet colour is independent of glyph and first run colour', () => {
  it('inherits a colour-only list style when the paragraph supplies the glyph', () => {
    const p = paragraph(brown, glyph)
    expect(p.bullet?.color).toBe('#9E3611')
    expect(p.runs[0].color).toBe('#FF0000')
    expect(generateParagraphXml(p)).toContain(brown)
  })
  it('keeps a paragraph colour override even when its glyph is inherited', () => {
    const p = paragraph(brown + glyph, '<a:buClr><a:srgbClr val="008000"/></a:buClr>')
    expect(p.bullet?.color).toBe('#008000')
    expect(generateParagraphXml(p)).toContain('val="008000"')
  })
  it.each([glyph, ''])('preserves explicit follow-text semantics on rebuild (%s)', (localGlyph) => {
    const p = paragraph(brown + glyph, '<a:buClrTx/>' + localGlyph)
    expect(p.bullet?.color).toBeUndefined()
    expect(p.bullet?.colorFollowsText).toBe(true)
    const xml = generateParagraphXml(p)
    expect(xml).toContain('<a:buClrTx/>')
    expect(xml).not.toContain('<a:buClr>')
  })
  it('inherits follow-text independently and allows an explicit colour to override it', () => {
    expect(paragraph('<a:buClrTx/>' + glyph, glyph).bullet?.colorFollowsText).toBe(true)
    expect(paragraph('<a:buClrTx/>' + glyph, brown).bullet?.color).toBe('#9E3611')
  })
})
