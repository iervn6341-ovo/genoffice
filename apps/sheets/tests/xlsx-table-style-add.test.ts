import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetTableAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { TableStyleAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-table-style-add'
import { buildEditFixture } from './fixture-builder'

const LEDGER: TableStyleAddition = {
  name: 'Ledger',
  elements: [
    { type: 'wholeTable', line: 'thin', lineColor: '#808080', edges: ['top', 'bottom'] },
    { type: 'headerRow', fill: '#1F3864', font: '#FFFFFF', bold: true },
    { type: 'firstRowStripe', fill: '#E2EFDA' },
  ],
}

const table = (style: string): SheetTableAddition => ({
  sheetName: 'Data',
  area: { startRow: 0, startColumn: 3, endRow: 3, endColumn: 4 },
  name: 'Users',
  columnNames: ['A', 'B'],
  style,
  bandedRows: true,
})

async function plan(
  styles: TableStyleAddition[],
  tables: SheetTableAddition[] = [],
  buffer?: Buffer,
) {
  const source = await createBufferEntrySource(buffer ?? (await buildEditFixture()))
  return planCellEditsToXlsx(
    source,
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    tables,
    [],
    [],
    [],
    [],
    [],
    [],
    null,
    null,
    [],
    [],
    styles,
  )
}

describe('custom table styles in styles.xml', () => {
  it('writes one dxf per formatted element and a tableStyle that names them', async () => {
    const result = await plan([LEDGER], [table('Ledger')])
    const styles = result.replaced.get('xl/styles.xml')!
    expect(styles).toContain('<tableStyles count="1"')
    const style = /<tableStyle name="Ledger" pivot="0" count="3">([\s\S]*?)<\/tableStyle>/.exec(
      styles,
    )
    expect(style).not.toBeNull()
    const elements = [...style![1]!.matchAll(/<tableStyleElement type="(\w+)" dxfId="(\d+)"\/>/g)]
    expect(elements.map((match) => match[1])).toEqual(['wholeTable', 'headerRow', 'firstRowStripe'])
    // every dxfId names a dxf in <dxfs>
    const dxfs = [...(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(styles)?.[1] ?? '').matchAll(/<dxf>/g)]
    for (const match of elements) expect(Number(match[2])).toBeLessThan(dxfs.length)
    expect(styles).toContain('<bgColor rgb="FF1F3864"/>')
    expect(styles).toMatch(/<dxfs count="\d+">/)
    // the table names the style
    expect(result.added.get('xl/tables/table1.xml')).toContain('<tableStyleInfo name="Ledger"')
  })

  it('puts tableStyles after dxfs, as the schema orders them', async () => {
    const styles = (await plan([LEDGER])).replaced.get('xl/styles.xml')!
    expect(styles.indexOf('</dxfs>')).toBeGreaterThan(-1)
    expect(styles.indexOf('<tableStyles')).toBeGreaterThan(styles.indexOf('</dxfs>'))
  })

  it('leaves a style the file already has under the name alone (a second save does not write it twice)', async () => {
    const first = (await plan([LEDGER])).replaced.get('xl/styles.xml')!
    // a stylesheet that already holds Ledger
    const source = await createBufferEntrySource(await buildEditFixture())
    void source
    const again = await planAgainst(first, [LEDGER])
    expect(again.replaced.get('xl/styles.xml') ?? first).toBe(first)
    expect([...first.matchAll(/<tableStyle name="Ledger"/g)]).toHaveLength(1)
  })

  it('refuses a built-in name, a color that is not #RRGGBB and an unknown element', async () => {
    await expect(plan([{ name: 'TableStyleMedium2', elements: [] }])).rejects.toThrow(
      /one of Excel's own table styles/,
    )
    await expect(
      plan([{ name: 'X', elements: [{ type: 'headerRow', fill: 'red' }] }]),
    ).rejects.toThrow(/not a #RRGGBB color/)
    await expect(
      plan([{ name: 'X', elements: [{ type: 'nope' as never, fill: '#FFFFFF' }] }]),
    ).rejects.toThrow(/not a table style element/)
  })

  it('leaves the stylesheet untouched when there is nothing to add', async () => {
    const result = await plan([])
    expect(result.replaced.get('xl/styles.xml')).toBeUndefined()
  })
})

/// a plan over a workbook whose styles.xml is `stylesXml`
async function planAgainst(stylesXml: string, styles: TableStyleAddition[]) {
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(await buildEditFixture())
  zip.file('xl/styles.xml', stylesXml)
  const buffer = await zip.generateAsync({ type: 'nodebuffer' })
  return plan(styles, [], buffer)
}
