import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '../src/gateway/csv-import'
import { applyCellEditsToXlsx, type CellEdit } from '../src/gateway/xlsx-gateway'

const SHEET_XML = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="D8:F11"/>
<cols><col min="4" max="5" style="2"/><col min="5" max="5" style="3"/></cols>
<sheetData>
  <row r="8" s="1" customFormat="1">
    <c r="D8" s="0"/>
    <c r="E8" t="inlineStr"><is><t>ONE&#10;TWO&#10;THREE</t></is></c>
    <c r="F8" s="0"><f>SUM(A1:A2)</f><v>5</v></c>
  </row>
  <row r="9">
    <c r="D9" t="inlineStr"><is><t>LEFT&#10;RIGHT</t></is></c>
    <c r="E9"><f>SUM(A1:A2)</f><v>5</v></c>
    <c r="F9"><v>12</v></c>
  </row>
  <row r="10" s="1"><c r="E10"><v>7</v></c></row>
  <row r="11" s="0" customFormat="1"><c r="E11"><v>8</v></c></row>
</sheetData></worksheet>`

async function fixture(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await blankXlsxBuffer())
  const styles = (await zip.file('xl/styles.xml')?.async('text')) ?? ''
  zip.file(
    'xl/styles.xml',
    styles.replace(
      /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,
      `<cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" horizontal="center" vertical="justify"/></xf>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" horizontal="right" vertical="distributed"/></xf>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" horizontal="left" vertical="top"/></xf>
      </cellXfs>`,
    ),
  )
  zip.file('xl/worksheets/sheet1.xml', SHEET_XML)
  return zip.generateAsync({ type: 'nodebuffer' })
}

function cell(xml: string, address: string): string {
  const opening = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>`).exec(xml)
  if (!opening) throw new Error(`Missing ${address}`)
  if (opening[0].endsWith('/>')) return opening[0]
  const start = opening.index
  const end = xml.indexOf('</c>', start + opening[0].length)
  if (end < 0) throw new Error(`Unclosed ${address}`)
  return xml.slice(start, end + '</c>'.length)
}

function xf(xml: string, index: number): string {
  const block = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? ''
  const entries = [...block.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)]
  const found = entries[index]?.[0]
  if (!found) throw new Error(`Missing XF ${index}`)
  return found
}

function styledXf(sheetXml: string, stylesXml: string, address: string): string {
  const index = Number(/\bs="(\d+)"/.exec(cell(sheetXml, address))?.[1])
  if (!Number.isInteger(index)) throw new Error(`Missing style on ${address}`)
  return xf(stylesXml, index)
}

describe('style-only edits of inherited cell formats', () => {
  it('uses cell > custom-format row > last column span > XF 0 without changing cell content', async () => {
    const source = await fixture()
    const edits: CellEdit[] = [
      [7, 3], // D8: explicit blank s="0" blocks the row.
      [7, 4], // E8: row 8 overrides both column spans.
      [7, 5], // F8: explicit s="0" overrides the row.
      [8, 3], // D9: first column span.
      [8, 4], // E9: last overlapping column span.
      [8, 5], // F9: base XF 0.
      [9, 4], // E10: row s without customFormat is ignored.
      [10, 4], // E11: customFormat row with explicit XF0 blocks column.
    ].map(([row, column]) => ({
      sheetName: 'Sheet1',
      row: row!,
      column: column!,
      writeValue: false,
      cell: { value: null },
      style: { verticalAlignment: 'distributed' },
    }))
    const mutation = await applyCellEditsToXlsx(source, edits)
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    const styles = (await zip.file('xl/styles.xml')?.async('text')) ?? ''

    expect(styledXf(sheet, styles, 'E8')).toContain('wrapText="1"')
    expect(styledXf(sheet, styles, 'E8')).toContain('horizontal="center"')
    expect(styledXf(sheet, styles, 'D9')).toContain('horizontal="right"')
    expect(styledXf(sheet, styles, 'E9')).toContain('horizontal="left"')
    expect(styledXf(sheet, styles, 'E10')).toContain('horizontal="left"')
    for (const address of ['D8', 'F8', 'F9', 'E11']) {
      expect(styledXf(sheet, styles, address)).not.toContain('wrapText="1"')
      expect(styledXf(sheet, styles, address)).not.toContain('horizontal=')
    }
    for (const address of ['D8', 'E8', 'D9', 'E9', 'F8', 'F9', 'E10', 'E11']) {
      expect(styledXf(sheet, styles, address)).toContain('vertical="distributed"')
      expect(cell(sheet, address).replace(/\s+s="\d+"/, '')).toBe(
        cell(SHEET_XML, address).replace(/\s+s="\d+"/, ''),
      )
    }
  })

  it('keeps explicit Reset style semantics on an XF-less row-inherited cell', async () => {
    const mutation = await applyCellEditsToXlsx(await fixture(), [
      {
        sheetName: 'Sheet1',
        row: 7,
        column: 4,
        writeValue: false,
        cell: { value: null },
        styleReset: true,
        style: { verticalAlignment: 'top' },
      },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    const styles = (await zip.file('xl/styles.xml')?.async('text')) ?? ''
    expect(styledXf(sheet, styles, 'E8')).toContain('vertical="top"')
    expect(styledXf(sheet, styles, 'E8')).not.toContain('wrapText="1"')
    expect(cell(sheet, 'E8')).toContain('ONE&#10;TWO&#10;THREE')
  })
})
