import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { generateParagraphXml, patchTextElementXml } from '../src/generate'
import type { TableElement, TextElement } from '../src/types'

const math =
  '<a14:m><m:oMath><m:sSub><m:e><m:r><a:rPr sz="1100"/><m:t>T</m:t></m:r></m:e><m:sub><m:r><m:t>k,com</m:t></m:r></m:sub></m:sSub></m:oMath></a14:m>'
const body = (content: string) => `<a:txBody><a:bodyPr/><a:p>${content}</a:p></a:txBody>`
const frame = (content: string, fill = '') =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Symbols"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="400000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000000"/></a:tblGrid><a:tr h="400000"><a:tc>${body(content)}<a:tcPr>${fill}</a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
const preview =
  '<a:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect t="-200000" b="-300000"/></a:stretch></a:blipFill>'

describe('PowerPoint table equations', () => {
  it('retains direct OMML and attaches the matching fallback image only for display', () => {
    const slideXml = `<p:sld><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="a14">${frame(math)}</mc:Choice><mc:Fallback>${frame('', preview)}</mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>`
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml,
      ctx: { mediaRels: new Map([['rId3', 'ppt/media/equation.png']]) },
    })
    const cell = (slide.elements[0] as TableElement).rows[0]![0]!
    expect(cell.mathPreview).toMatchObject({
      type: 'image',
      mediaRef: 'ppt/media/equation.png',
      fillRect: { t: -2, b: -3 },
    })
    expect(cell.fill).toBeUndefined()
    const p = cell.text!.paragraphs[0]!
    expect(p.runs[0]).toMatchObject({ text: 'Tk,com', rawXml: math })
    expect(generateParagraphXml(p)).toContain(math)
  })
  it('keeps native equation XML byte-for-byte when surrounding text is edited', () => {
    const sp = `<p:sp><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Before</a:t></a:r>${math}<a:r><a:rPr/><a:t>After</a:t></a:r></a:p></p:txBody></p:sp>`
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: `<p:sld><p:cSld><p:spTree>${sp}</p:spTree></p:cSld></p:sld>`,
      ctx: {},
    })
    const el = slide.elements[0] as TextElement
    expect(patchTextElementXml(el, sp)).toBe(sp)
    el.text!.paragraphs[0]!.runs[0]!.text = 'Edited'
    expect(patchTextElementXml(el, sp)).toBe(sp.replace('Before', 'Edited'))
  })
})
