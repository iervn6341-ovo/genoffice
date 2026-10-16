import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, type CellEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

const ORIGINAL_XFS =
  '<cellStyleXfs count="2"><xf/><xf applyAlignment="1"><alignment horizontal="center" vertical="justify" wrapText="1" indent="3"/></xf></cellStyleXfs><cellXfs count="4"><xf xfId="0"/><xf xfId="1" applyAlignment="0"/><xf xfId="1" applyAlignment="1"/><xf xfId="1" applyAlignment="1"><alignment horizontal="left"/></xf></cellXfs>'

async function fixture(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file(
    'xl/styles.xml',
    styles.replace(
      '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>',
      ORIGINAL_XFS,
    ),
  )
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace(
      '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="C1"><v>5</v></c></row>',
      '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" s="2"><v>8</v></c><c r="C1" s="3"><v>5</v></c><c r="D1"><v>9</v></c></row>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function parts(buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer)
  return {
    sheet: await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
    styles: await zip.file('xl/styles.xml')!.async('string'),
  }
}

function cell(xml: string, address: string): string {
  const match = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`).exec(xml)
  if (!match) throw new Error(`Missing ${address}`)
  return match[0]
}

function edit(column: number, changes: Pick<CellEdit, 'style' | 'styleReference'>): CellEdit {
  return {
    sheetName: 'Data',
    row: 0,
    column,
    writeValue: false,
    cell: { value: null },
    ...changes,
  }
}

describe('cross-save cell style history', () => {
  it('restores exact named-style XF references and explicit-style absence without touching formulas', async () => {
    const original = await fixture()
    const edited = await applyCellEditsToXlsx(
      original,
      [0, 1, 2, 3].map((column) => edit(column, { style: { verticalAlignment: 'distributed' } })),
    )
    const restored = await applyCellEditsToXlsx(edited.buffer, [
      edit(0, { styleReference: 1 }),
      edit(1, { styleReference: 2 }),
      edit(2, { styleReference: 3 }),
      edit(3, { styleReference: null }),
    ])
    const first = await parts(original)
    const after = await parts(restored.buffer)
    for (const address of ['A1', 'B1', 'C1', 'D1', 'B3']) {
      expect(cell(after.sheet, address)).toBe(cell(first.sheet, address))
    }
    expect(after.styles).toContain(ORIGINAL_XFS.slice(0, ORIGINAL_XFS.indexOf('<cellXfs')))
    expect(cell(after.sheet, 'B3')).toContain('<f>SUM(C1:C2)</f><v>5</v>')

    const redone = await applyCellEditsToXlsx(restored.buffer, [
      edit(0, { styleReference: 1, style: { verticalAlignment: 'distributed' } }),
      edit(1, { styleReference: 2, style: { verticalAlignment: 'distributed' } }),
      edit(2, { styleReference: 3, style: { verticalAlignment: 'distributed' } }),
    ])
    const redoParts = await parts(redone.buffer)
    expect(redoParts.sheet).toContain('<f>SUM(C1:C2)</f><v>5</v>')
    expect(cell(redoParts.sheet, 'A1')).not.toContain('s="1"')
    expect(cell(redoParts.sheet, 'B1')).not.toContain('s="2"')
    expect(cell(redoParts.sheet, 'C1')).not.toContain('s="3"')
  })
})
