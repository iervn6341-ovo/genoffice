import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { foreignFromExcelEmbed, parseExcelLink } from '../src/renderer/excel-embed'

const embed = new Uint8Array(
  readFileSync(join(process.cwd(), 'tests/fixtures/excel-mac-embed.bin')),
)
const linkText = readFileSync(join(process.cwd(), 'tests/fixtures/excel-mac-link.bin')).toString(
  'latin1',
)

describe('Excel link address', () => {
  it('reads the sheet and the R1C1 range Excel names', () => {
    expect(parseExcelLink(linkText)).toEqual({
      sheetName: 'Sheet1',
      range: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 5 },
    })
    expect(parseExcelLink('Excel\0/x/[Book.xlsx]My Data\0R5C3\0')).toEqual({
      sheetName: 'My Data',
      range: { startRow: 4, startColumn: 2, endRow: 4, endColumn: 2 },
    })
    expect(parseExcelLink('nothing')).toBeNull()
  })
})

describe('Excel embedded workbook (real Excel for Mac data)', () => {
  const asStyle = (style: unknown): unknown => style
  it('gives formulas, exact values, styles and merges that the HTML flavor lacks', async () => {
    const link = parseExcelLink(linkText)!
    const copy = (await foreignFromExcelEmbed(embed, link, { toStyle: asStyle }))!
    expect(copy.origin).toBe('excel')
    expect(copy.literalFormulas).toBe(false)
    expect(copy.copied.rows).toEqual([0, 1, 2, 3])
    expect(copy.copied.columns).toEqual([0, 1, 2, 3, 4, 5])
    const cell = (i: number, j: number) => copy.copied.cells[i]![j]!
    expect(cell(0, 0)).toMatchObject({ v: 'Item', t: 1 })
    expect(cell(1, 1)).toMatchObject({ v: 1234.5, t: 2 })
    expect(cell(1, 2)).toMatchObject({ v: 0.256, t: 2 })
    expect(cell(1, 5)).toMatchObject({ f: '=B2*E2', v: 122832.75, t: 2 })
    expect(cell(2, 5)).toMatchObject({ f: '=SUM(B2:B3)' })
    expect(cell(3, 3)).toMatchObject({ f: '=Data!A2*2' })
    expect(cell(3, 4)).toMatchObject({ f: '=Rate*B3' })
    expect(cell(2, 1)).toMatchObject({ v: 1, t: 3 })
    expect(cell(2, 0).v).toBe('line1\nline2')
    expect(copy.merges).toEqual([{ startRow: 3, endRow: 3, startColumn: 0, endColumn: 1 }])
    expect(copy.rowHeights[2]).toBeCloseTo(53.33, 1)
  })

  it('reads the real number formats, fonts, fills and borders from styles.xml', async () => {
    const link = parseExcelLink(linkText)!
    const copy = (await foreignFromExcelEmbed(embed, link, { toStyle: asStyle }))!
    const style = (i: number, j: number) => copy.copied.cells[i]![j]!.s as Record<string, unknown>
    expect(style(1, 1)).toMatchObject({ numberFormat: '#,##0.00' })
    expect(style(1, 2)).toMatchObject({ numberFormat: '0.00%' })
    expect(style(1, 3)).toMatchObject({ numberFormat: 'yyyy\\-mm\\-dd' })
    expect(style(1, 4)).toMatchObject({ numberFormat: '[$$-409]#,##0.00' })
    expect(style(0, 0)).toMatchObject({
      fontFamily: 'Arial',
      fontSize: 12,
      bold: true,
      fontColor: '#FFFFFF',
      fillColor: '#FFC000',
      horizontalAlignment: 'center',
    })
    expect(style(1, 0)).toMatchObject({
      italic: true,
      fontFamily: 'Times New Roman',
      fontColor: '#FF0000',
    })
    expect(style(2, 0)).toMatchObject({
      wrapText: true,
      borderTop: { style: 'thin' },
      borderLeft: { style: 'thin' },
    })
    expect(style(2, 2)).toMatchObject({ underline: true, strikethrough: true })
    expect(style(2, 3)).toMatchObject({ borderBottom: { style: 'dashDot', color: '#0000FF' } })
    expect(style(2, 4)).toMatchObject({ borderLeft: { style: 'medium', color: '#FF0000' } })
    expect(style(3, 0)).toMatchObject({ fillColor: '#00B050', verticalAlignment: 'center' })
  })

  it('refuses data that is not an OLE file with a package', async () => {
    const link = parseExcelLink(linkText)!
    expect(
      await foreignFromExcelEmbed(new Uint8Array([1, 2, 3]), link, { toStyle: (s) => s }),
    ).toBeNull()
  })
})
