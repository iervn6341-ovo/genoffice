/**
 * HTML from other programs through the editor's paste pipeline: Word for Mac
 * (real clipboard captures in tests/fixtures/clipboard), Excel-style class
 * rules, web pages. Checks the model AND the saved docx, so a fill or border
 * that parses but never reaches the file is caught.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { cleanPastedHtml } from '../src/renderer/editor/paste-clean'
import {
  cssColorToHex,
  foreignBorders,
  lengthToPt,
  parseBorderValue,
  parseDeclarations,
} from '../src/renderer/editor/foreign-style'
import { blocks, caretAtEnd, openDocx, paste } from './helpers/clip-harness'

const fixture = (name: string) => readFileSync(join(__dirname, 'fixtures/clipboard', name), 'utf8')
const blankDoc = async () =>
  openDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>B doc</w:t></w:r></w:p>' }))

describe('css values', () => {
  it('reads color names, short hex and rgb()', () => {
    expect(cssColorToHex('red')).toBe('FF0000')
    expect(cssColorToHex('Purple')).toBe('800080')
    expect(cssColorToHex('#f80')).toBe('FF8800')
    expect(cssColorToHex('rgb(31, 78, 121)')).toBe('1F4E79')
    expect(cssColorToHex('rgba(0, 0, 0, 0)')).toBeNull()
    expect(cssColorToHex('windowtext')).toBeNull()
  })

  it('converts lengths to points', () => {
    expect(lengthToPt('1.5pt')).toBe(1.5)
    expect(lengthToPt('16px')).toBe(12)
    expect(lengthToPt('1in')).toBe(72)
    expect(lengthToPt('50%')).toBeNull()
  })

  it('parses borders in either token order, and none', () => {
    expect(parseBorderValue('solid red 1.5pt')).toEqual({
      style: 'single',
      szEighths: 12,
      color: 'FF0000',
    })
    expect(parseBorderValue('1px dashed #0000FF')).toEqual({
      style: 'dashed',
      szEighths: 6,
      color: '0000FF',
    })
    expect(parseBorderValue('double #00AA00 2.25pt')).toMatchObject({
      style: 'double',
      szEighths: 18,
    })
    expect(parseBorderValue('none')).toBeNull()
  })

  it('applies a shorthand border then a longhand override in order', () => {
    const el = document.createElement('td')
    el.setAttribute('style', 'border:solid red 1.5pt;border-left:none;border-bottom-color:blue')
    const b = foreignBorders(el)
    expect(b.top).toMatchObject({ style: 'single', color: 'FF0000' })
    expect(b.left).toBe('none')
    expect(b.bottom).toMatchObject({ color: '0000FF' })
  })

  it('keeps later declarations over earlier ones', () => {
    expect([...parseDeclarations('color:red;color:blue;x:url(a;b)')]).toEqual([
      ['color', 'blue'],
      ['x', 'url(a;b)'],
    ])
  })
})

describe('Word for Mac table → Docs', () => {
  it('keeps fills, per-cell borders, spans, alignment and row height', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    const html = fixture('word-mac-table.html')
    paste(dst.editor, { html, text: '' })
    const table = blocks(dst.editor).find((b) => b.type === 'docTable')!
    const rows = table.content!
    expect(rows[0].attrs?.heightTwips).toBe(600)
    const [merged, vm] = rows[0].content!
    expect(merged.attrs).toMatchObject({ colspan: 2, fill: 'FFFF00' })
    expect(merged.attrs?.borders).toMatchObject({
      top: { style: 'single', color: 'FF0000' },
      bottom: { style: 'dashed', color: '0000FF' },
    })
    expect(vm.attrs).toMatchObject({ rowspan: 2, fill: '00FF00' })
    expect(vm.attrs?.borders).toMatchObject({ left: { style: 'nil' }, top: { color: 'FF0000' } })
    const a2 = rows[1].content![0]
    expect(a2.content![0].attrs).toMatchObject({ align: 'right' })
    expect(a2.attrs?.borders).toMatchObject({ bottom: { style: 'double', color: '00AA00' } })
    // 270pt merged cell = 90pt + 180pt columns → 120px + 240px
    expect(merged.attrs?.colwidth).toEqual([120, 240])
  })

  it('saves the fills and borders into the docx', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    paste(dst.editor, { html: fixture('word-mac-table.html'), text: '' })
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('w:fill="FFFF00"')
    expect(documentXml).toContain('w:fill="00FF00"')
    expect(documentXml).toMatch(/<w:top w:val="single"[^>]*w:color="FF0000"/)
    expect(documentXml).toMatch(/w:val="dashed"[^>]*w:color="0000FF"/)
    expect(documentXml).toContain('w:gridSpan w:val="2"')
    expect(documentXml).toContain('<w:vMerge w:val="restart"/>')
  })
})

describe('Word for Mac paragraphs → Docs', () => {
  it('resolves class rules from the style block: highlight, character color, Courier', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    paste(dst.editor, { html: fixture('word-mac-paras.html'), text: '' })
    const para = blocks(dst.editor).find((b) => JSON.stringify(b).includes('red14'))!
    const text = JSON.stringify(para)
    expect(text).toContain('"color":"FF0000"') // color:red (a name, not hex)
    expect(text).toContain('"highlight":"yellow"')
    expect(text).toContain('Courier New')
  })

  it('keeps the shading of a styled paragraph and its alignment', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    paste(dst.editor, { html: fixture('word-mac-paras.html'), text: '' })
    const callout = blocks(dst.editor).find((b) => JSON.stringify(b).includes('Callout paragraph'))!
    expect(callout.attrs).toMatchObject({ align: 'center', shadingFill: 'FFF2CC' })
  })

  it('turns mso-list paragraphs into list items at their level', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    paste(dst.editor, { html: fixture('word-mac-paras.html'), text: '' })
    const items = blocks(dst.editor).filter((b) => b.type === 'docListItem')
    expect(items.map((i) => i.attrs?.ilvl)).toEqual([0, 0, 1])
    // the "(a)" / "i." marker span is not part of the text
    expect(items.map((i) => i.content![0].text)).toEqual(['first (a)', 'second (b)', 'sub i.'])
    expect(items.every((i) => i.attrs?.kind === 'ordered')).toBe(true)
  })

  it('keeps the hyperlink', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    paste(dst.editor, { html: fixture('word-mac-paras.html'), text: '' })
    expect(JSON.stringify(blocks(dst.editor))).toContain('https://example.com/')
  })
})

describe('Excel-style class rules → Docs', () => {
  const EXCEL =
    '<html><head><style>.xl65{background:#FFC000;font-weight:700;border:.5pt solid black}' +
    '.xl66{color:#FF0000;text-align:center}</style></head><body><table>' +
    '<tr><td class=xl65>Item</td><td class=xl65>Qty</td></tr><tr><td>a</td><td class=xl66>3</td></tr></table></body></html>'

  it('folds the class rules into the cells', async () => {
    const dst = await blankDoc()
    caretAtEnd(dst.editor)
    paste(dst.editor, { html: EXCEL, text: 'Item\tQty\na\t3' })
    const table = blocks(dst.editor).find((b) => b.type === 'docTable')!
    const head = table.content![0].content![0]
    expect(head.attrs).toMatchObject({ fill: 'FFC000' })
    expect(head.attrs?.borders).toMatchObject({ top: { style: 'single', color: '000000' } })
    expect(table.content![1].content![1].attrs).toMatchObject({ align: 'center' })
  })
})

describe('cleanPastedHtml', () => {
  it('leaves GenOffice payload html untouched (a one-cell table is not unwrapped)', () => {
    const html =
      '<table data-pm-slice="0 0 []" data-genoffice-slice="e30="><tbody><tr><td><p>x</p></td></tr></tbody></table>'
    expect(cleanPastedHtml(html)).toBe(html)
  })
})
