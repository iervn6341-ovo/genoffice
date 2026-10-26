// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import type { IStyleData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  buildClipboardHtml,
  cellCss,
  msoNumberFormatOf,
  type HtmlBlockCell,
} from '../src/renderer/clipboard-html'
import {
  cssColor,
  cssLengthToPx,
  foreignFromExcelHtml,
  looksLikeExcelHtml,
  msoNumberFormat,
  parseCssDeclarations,
  parseStyleSheet,
  styleFromDeclarations,
} from '../src/renderer/clipboard-html-import'
import { formulaForExcelClipboard, nameReferences } from '../src/renderer/clipboard-formula'
import { withFullDisplay } from '../src/renderer/clipboard-tsv'

const excelHtml = readFileSync(join(__dirname, 'fixtures/excel-mac-clip.html'), 'utf8')

describe('mso-number-format', () => {
  it('writes patterns the way Excel escapes them and reads them back', () => {
    for (const pattern of [
      '#,##0.00',
      '0.00%',
      'yyyy-mm-dd',
      '[$$-409]#,##0.00',
      '0.00;[Red]-0.00',
      '"Qty" 0',
      '@',
    ]) {
      const written = msoNumberFormatOf(pattern)
      expect(msoNumberFormat(written)).toBe(pattern)
    }
    expect(msoNumberFormatOf('#,##0.00')).toBe('"\\#\\,\\#\\#0\\.00"')
  })

  it('maps Excel named formats', () => {
    expect(msoNumberFormat('Standard')).toBe('#,##0.00')
    expect(msoNumberFormat('Percent')).toBe('0.00%')
    expect(msoNumberFormat('Short Date')).toBe('m/d/yyyy')
    expect(msoNumberFormat('General')).toBeNull()
    expect(msoNumberFormat('"\\@"')).toBe('@')
  })
})

describe('css helpers', () => {
  it('parses declarations with quotes and parentheses', () => {
    expect(
      parseCssDeclarations(
        'font-family:"Times New Roman";color:rgb(1,2,3);mso-number-format:"\\#\\,\\#"',
      ),
    ).toEqual({
      'font-family': '"Times New Roman"',
      color: 'rgb(1,2,3)',
      'mso-number-format': '"\\#\\,\\#"',
    })
  })
  it('reads colors and lengths', () => {
    expect(cssColor('windowtext')).toBe('#000000')
    expect(cssColor('red')).toBe('#FF0000')
    expect(cssColor('#fc0')).toBe('#FFCC00')
    expect(cssColor('rgb(255, 192, 0)')).toBe('#FFC000')
    expect(cssColor('none')).toBeNull()
    expect(cssLengthToPx('.5pt')).toBeCloseTo(0.667, 2)
    expect(cssLengthToPx('120pt')).toBe(160)
  })
  it('parses a style block with comments', () => {
    const rules = parseStyleSheet('<!--table{a:b}\n.xl1\n\t{color:red;\n\tfont-weight:700;}\n-->')
    expect(rules.get('.xl1')).toEqual({ color: 'red', 'font-weight': '700' })
  })
})

describe('Excel markup to a foreign copy (real Excel for Mac output)', () => {
  const copy = foreignFromExcelHtml(excelHtml)!
  const cell = (i: number, j: number) => copy.copied.cells[i]![j]!
  const st = (i: number, j: number) => cell(i, j).s as IStyleData | null

  it('is recognized as Excel markup', () => {
    expect(looksLikeExcelHtml(excelHtml)).toBe(true)
    expect(looksLikeExcelHtml('<table><tr><td>a</td></tr></table>')).toBe(false)
    expect(foreignFromExcelHtml('<table><tr><td>a</td></tr></table>')).toBeNull()
    expect(copy.origin).toBe('excel')
    expect(copy.copied.rows).toHaveLength(4)
    expect(copy.copied.columns).toHaveLength(6)
  })

  it('keeps fonts, fills and alignment from the class rules', () => {
    expect(st(0, 0)).toMatchObject({
      ff: 'Arial',
      fs: 12,
      bl: 1,
      cl: { rgb: '#FFFFFF' },
      bg: { rgb: '#FFC000' },
      ht: 2,
    })
    expect(st(1, 0)).toMatchObject({ ff: 'Times New Roman', fs: 14, it: 1, cl: { rgb: '#FF0000' } })
    // the inline fill of a cell colored by conditional formatting
    expect(st(1, 1)).toMatchObject({ bg: { rgb: '#FF9999' } })
  })

  it('turns displayed numbers back into numbers with the cell number format', () => {
    expect(cell(1, 1)).toMatchObject({ v: 1234.5, t: 2, s: { n: { pattern: '#,##0.00' } } })
    expect(cell(1, 2)).toMatchObject({ v: 0.256, t: 2, s: { n: { pattern: '0.00%' } } })
    expect(st(1, 3)?.n?.pattern).toBe('yyyy\\-mm\\-dd')
    expect(typeof cell(1, 3).v).toBe('number')
    expect(cell(1, 4)).toMatchObject({ v: 99.5, s: { n: { pattern: '[$$-409]#,##0.00' } } })
    expect(cell(1, 5)).toMatchObject({ v: 122833, s: { n: { pattern: '#,##0' } } })
  })

  it('reads borders, wrap, merges, text decoration and line breaks', () => {
    expect(st(2, 0)?.bd).toMatchObject({ t: { s: 1 }, r: { s: 1 }, b: { s: 1 }, l: { s: 1 } })
    expect(st(2, 0)?.tb).toBe(3)
    expect(cell(2, 0).v).toBe('line1\nline2')
    expect(cell(2, 1)).toMatchObject({ v: 1, t: 3 })
    expect(st(2, 2)?.ul).toMatchObject({ s: 1 })
    expect(st(2, 2)?.st).toMatchObject({ s: 1 })
    expect(st(2, 3)?.bd?.b?.cl?.rgb).toBe('#0000FF')
    expect(st(2, 4)?.bd?.l).toMatchObject({ s: 8, cl: { rgb: '#FF0000' } })
    expect(copy.merges).toEqual([{ startRow: 3, endRow: 3, startColumn: 0, endColumn: 1 }])
    expect(st(3, 0)).toMatchObject({ bg: { rgb: '#00B050' }, ht: 2, vt: 2 })
  })

  it('reads sizes and leaves formulas as literal (Excel markup has none)', () => {
    expect(copy.columnWidths[0]).toBe(160)
    expect(copy.rowHeights[0]).toBeCloseTo(21.33, 1)
    expect(copy.literalFormulas).toBe(true)
  })
})

describe('the markup this app writes', () => {
  const cells: HtmlBlockCell[][] = [
    [
      {
        text: '1,234.50',
        t: 2,
        v: 1234.5,
        style: { n: { pattern: '#,##0.00' }, bg: { rgb: '#FFC000' }, bl: 1, ht: 2 },
      },
      { text: '122,808', t: 2, v: 122807.75, f: '=A1*2', style: { n: { pattern: '#,##0' } } },
    ],
    [
      {
        text: 'a\nb <&>',
        t: 1,
        v: 'a\nb <&>',
        style: { tb: 3, bd: { l: { s: 8, cl: { rgb: '#FF0000' } } } },
      },
      { text: '00123', t: 1, v: '00123', style: null },
    ],
    [
      { text: 'm', t: 1, v: 'm', style: null, colSpan: 2, rowSpan: 1 },
      { text: '', style: null, covered: true },
    ],
  ]
  const html = buildClipboardHtml({
    copyId: 'id1',
    payloadValue: 'QUJD',
    cells,
    columnWidths: [100, null],
    rowHeights: [20, 30, null],
  })

  it('carries the copy id, the payload and Excel typed values', () => {
    expect(html).toContain('data-copy-id="id1"')
    expect(html).toContain('data-genoffice-sheet="QUJD"')
    expect(html).toContain('x:num="1234.5"')
    expect(html).toContain('x:fmla="=A1*2"')
    expect(html).toContain('mso-number-format:"\\#\\,\\#\\#0\\.00"')
    expect(html).toContain('x:str="&#39;00123"'.replace('&#39;', "'"))
    expect(html).toContain("<br style='mso-data-placement:same-cell;'>")
    expect(html).toContain('a<br')
    expect(html).toContain('&lt;&amp;&gt;')
    expect(html).toContain('border-left:1.0pt solid #FF0000')
    expect(html).toContain('background:#FFC000;mso-pattern:black none')
    expect(html).toContain('colspan=2')
    expect((html.match(/<td /g) ?? []).length).toBe(5)
    expect(html).toContain('<col width=100')
  })

  it('reads back through the importer into the same cells', () => {
    const copy = foreignFromExcelHtml(html)!
    const c = copy.copied.cells
    expect(c[0]![0]).toMatchObject({
      v: 1234.5,
      t: 2,
      s: { n: { pattern: '#,##0.00' }, bl: 1, ht: 2, bg: { rgb: '#FFC000' } },
    })
    expect(c[1]![0]).toMatchObject({ v: 'a\nb <&>', t: 1 })
    expect((c[1]![0]!.s as IStyleData).tb).toBe(3)
    expect((c[1]![0]!.s as IStyleData).bd?.l).toMatchObject({ s: 8, cl: { rgb: '#FF0000' } })
    expect(c[1]![1]).toMatchObject({ v: '00123', t: 1 })
    expect(copy.merges).toEqual([{ startRow: 2, endRow: 2, startColumn: 0, endColumn: 1 }])
  })

  it('writes cell css for every look the model has', () => {
    const css = cellCss(
      {
        ff: 'Times New Roman',
        fs: 14,
        it: 1,
        ul: { s: 1 },
        st: { s: 1 },
        cl: { rgb: '#FF0000' },
        vt: 2,
        ht: 3,
      },
      { hasNumber: false },
    )
    expect(css).toContain('font-family:"Times New Roman"')
    expect(css).toContain('font-size:14pt')
    expect(css).toContain('font-style:italic')
    expect(css).toContain('text-decoration:underline line-through')
    expect(css).toContain('vertical-align:middle')
    expect(css).toContain('text-align:right')
  })
})

describe('style declarations', () => {
  it('reads border shorthands and none sides', () => {
    const style = styleFromDeclarations({ border: '.5pt solid black', 'border-top': 'none' })
    expect(style?.bd?.t).toBeUndefined()
    expect(style?.bd?.b).toMatchObject({ s: 1 })
  })
})

describe('formulas for Excel', () => {
  it('moves references so a block pasted anywhere keeps its relations', () => {
    expect(formulaForExcelClipboard('=B3*C3', 2, 0)).toBe('=B1*C1')
    expect(formulaForExcelClipboard('=$B$3+C3', 2, 1)).toBe('=$B$3+B1')
    // a reference above or left of the block cannot be written relative to A1
    expect(formulaForExcelClipboard('=A1+1', 5, 5)).toBeNull()
  })
  it('finds defined names but not cells, functions or text', () => {
    expect(nameReferences('=Rate*B3')).toEqual(['RATE'])
    expect(nameReferences('=SUM(A1:B2)+IF(TRUE,"x y",1)')).toEqual([])
    expect(nameReferences('=SUM(A:C)')).toEqual([])
    expect(nameReferences('=tax_rate+B2*Total.Amount')).toEqual(['TAX_RATE', 'TOTAL.AMOUNT'])
  })
})

describe('full display text', () => {
  it('replaces the #### of a narrow column by the formatted text', () => {
    expect(withFullDisplay({ v: 46000, displayV: '#######' }, 'yyyy-mm-dd')?.displayV).toBe(
      '2025-12-09',
    )
    expect(withFullDisplay({ v: 1234.5, displayV: '###' }, null)?.displayV).toBe('1234.5')
    expect(withFullDisplay({ v: 5, displayV: '5' }, '0.00')?.displayV).toBe('5')
    expect(withFullDisplay({ v: 'text', displayV: '##' }, null)?.displayV).toBe('##')
  })
})
