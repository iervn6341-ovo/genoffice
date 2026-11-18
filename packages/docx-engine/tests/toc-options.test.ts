import { describe, expect, it } from 'vitest'
import {
  DEFAULT_TOC_OPTIONS,
  generateTocFieldXml,
  parseTocInstruction,
  patchTocCachedPageNumber,
  tocInstruction,
  tocInstructionFromXml,
  tocOptionsFromField,
} from '../src'

describe('native TOC field options', () => {
  it('writes the chosen native switches and cache formatting', () => {
    const options = {
      ...DEFAULT_TOC_OPTIONS,
      levels: 4,
      showPageNumbers: false,
      useHyperlinks: false,
      useTcFields: true,
      styleLevels: { CustomTitle: 2 },
      format: 'modern' as const,
    }
    const xml = generateTocFieldXml([{ level: 2, text: 'Chapter & Title', pageNo: 8 }], options)[0]!
    expect(tocInstructionFromXml(xml)).toBe(tocInstruction(options))
    expect(xml).toContain('\\t "CustomTitle,2"')
    expect(xml).toContain('\\f')
    expect(xml).toContain('\\n')
    expect(xml).not.toContain('<w:tab/>')
    expect(xml).toContain('Chapter &amp; Title')
    expect(xml).toContain('<w:color w:val="4472C4"/>')
    expect(tocOptionsFromField(tocInstruction(options), xml)).toMatchObject({
      showPageNumbers: false,
      format: 'modern',
      styleLevels: { CustomTitle: 2 },
    })
  })

  it('reads supported imported instructions and refuses switches we cannot update safely', () => {
    const options = parseTocInstruction(' TOC \\o "1-4" \\t "MyStyle,2" \\f \\h \\z \\u ')
    expect(options).toMatchObject({
      levels: 4,
      useHeadingStyles: true,
      useTcFields: true,
      useHyperlinks: true,
      styleLevels: { MyStyle: 2 },
    })
    expect(parseTocInstruction(' TOC \\o "2-4" \\h ')).toBeNull()
    expect(parseTocInstruction(' TOC \\b Chapter1 \\o "1-3" ')).toBeNull()
    expect(parseTocInstruction(' TOC \\n "3-4" ')).toBeNull()
    const mixed =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> PAGEREF _Toc42 \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    expect(tocInstructionFromXml(mixed)).toBe(' TOC \\o "1-3" ')
  })

  it('changes only the cached page run and preserves nested fields and formatting', () => {
    const xml =
      '<w:p><w:r><w:t>Title</w:t></w:r><w:r><w:tab/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>7</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    expect(patchTocCachedPageNumber(xml, '7', '12')).toBe(
      xml.replace('<w:t>7</w:t>', '<w:t>12</w:t>'),
    )
    expect(patchTocCachedPageNumber(xml, '9', '12')).toBeNull()
    expect(patchTocCachedPageNumber(xml.replace('<w:tab/>', ''), '7', '12')).toBeNull()
  })

  it('supports caption switch \\c and hyperlinked entries with PAGEREF', () => {
    const parsed = parseTocInstruction(' TOC \\h \\z \\c "圖表" ')
    expect(parsed).toMatchObject({
      captionLabel: '圖表',
      useHyperlinks: true,
      useHeadingStyles: false,
    })
    expect(tocInstruction(parsed!)).toBe(' TOC \\c "圖表" \\h \\z ')

    const options = {
      ...DEFAULT_TOC_OPTIONS,
      useHyperlinks: true,
    }
    const xml = generateTocFieldXml(
      [{ level: 1, text: '緒論', num: '第1章', pageNo: 1, anchor: '_Toc245137332' }],
      options,
    )[0]!

    expect(xml).toContain('<w:hyperlink w:anchor="_Toc245137332" w:history="1">')
    expect(xml).toContain('PAGEREF _Toc245137332 \\h')
    expect(xml).toContain('第1章')
    expect(xml).toContain('緒論')
    expect(xml).toContain('<w:t>1</w:t>')
    expect(xml).toContain('</w:hyperlink>')
  })
})
