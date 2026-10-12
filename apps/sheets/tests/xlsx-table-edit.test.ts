import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  patchTableStyleInfo,
  tableInfoOf,
  tableKeyOf,
  type TableEdit,
} from '@genoffice/xlsx-gateway/gateway/xlsx-table-edit'
import { buildEditFixture } from './fixture-builder'

const TABLE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4" totalsRowShown="0">' +
  '<autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns>' +
  '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'

async function planEdits(edits: TableEdit[], table = TABLE) {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  zip.file('xl/tables/table1.xml', table)
  const source = await createBufferEntrySource(await zip.generateAsync({ type: 'nodebuffer' }))
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
    [],
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
    [],
    edits,
  )
}

const info = (xml: string): string => /<tableStyleInfo [^>]*\/>/.exec(xml)![0]

describe('Table Design on a table the file already has patches its tableStyleInfo only', () => {
  it('changes the style, and the options that are given, and keeps the rest as the file has them', () => {
    const out = patchTableStyleInfo(TABLE, {
      name: 'Sales',
      style: 'TableStyleDark9',
      firstColumn: true,
    })
    expect(info(out)).toBe(
      '<tableStyleInfo name="TableStyleDark9" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    )
    // the columns, the filter and the ref are untouched
    expect(out).toContain('<autoFilter ref="A1:C4"/>')
    expect(out).toContain('<tableColumn id="2" name="b"/>')
    expect(out).toContain('ref="A1:C4"')
  })

  it('turns options off as well as on, and leaves the style alone when none is given', () => {
    const out = patchTableStyleInfo(TABLE, {
      name: 'Sales',
      bandedRows: false,
      bandedColumns: true,
    })
    expect(info(out)).toBe(
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="1"/>',
    )
  })

  it('writes a custom style name, escaped', () => {
    expect(
      info(patchTableStyleInfo(TABLE, { name: 'Sales', style: 'My "Ledger" & Co' })),
    ).toContain('name="My &quot;Ledger&quot; &amp; Co"')
  })

  it('adds a tableStyleInfo to a table that has none, where the schema puts it (after tableColumns)', () => {
    const bare = TABLE.replace(/<tableStyleInfo [^>]*\/>/, '')
    const out = patchTableStyleInfo(bare, { name: 'Sales', style: 'TableStyleLight9' })
    expect(out.indexOf('<tableStyleInfo')).toBeGreaterThan(out.indexOf('</tableColumns>'))
    expect(out.indexOf('<tableStyleInfo')).toBeLessThan(out.indexOf('</table>'))
    expect(info(out)).toContain('name="TableStyleLight9"')
    // before an extLst
    const withExt = bare.replace('</table>', '<extLst><ext uri="x"/></extLst></table>')
    const outExt = patchTableStyleInfo(withExt, { name: 'Sales', style: 'TableStyleLight9' })
    expect(outExt.indexOf('<tableStyleInfo')).toBeLessThan(outExt.indexOf('<extLst'))
  })

  it('reads what the part says (for the options Table Design shows), by name in any case', () => {
    expect(tableInfoOf(TABLE, 'sales')).toEqual({
      style: 'TableStyleMedium2',
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    })
    expect(tableInfoOf(TABLE, 'Other')).toBeNull()
    expect(tableInfoOf(TABLE.replace(/<tableStyleInfo [^>]*\/>/, ''), 'Sales')).toEqual({
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: false,
      showColumnStripes: false,
    })
  })
})

describe('the save writes the table edits into the table part', () => {
  it('finds the table by its name and patches only its part', async () => {
    const plan = await planEdits([
      { name: 'SALES', style: 'TableStyleDark9', lastColumn: true, bandedRows: false },
    ])
    const part = plan.replaced.get('xl/tables/table1.xml')!
    expect(info(part)).toBe(
      '<tableStyleInfo name="TableStyleDark9" showFirstColumn="0" showLastColumn="1" showRowStripes="0" showColumnStripes="0"/>',
    )
    expect(plan.touchedEntries).toContain('xl/tables/table1.xml')
  })

  it('refuses a table the workbook does not have', async () => {
    await expect(planEdits([{ name: 'Nope', style: 'TableStyleDark9' }])).rejects.toThrow(
      /not in the workbook/,
    )
  })

  it('writes nothing when there are no edits', async () => {
    const plan = await planEdits([])
    expect(plan.replaced.get('xl/tables/table1.xml')).toBeUndefined()
  })
})

describe('the table part is read the way the file writes it', () => {
  it('patches a style info written out in full or with single quotes, once', () => {
    const open = TABLE.replace(
      /<tableStyleInfo [^>]*\/>/,
      "<tableStyleInfo name='TableStyleMedium2' showRowStripes='1'></tableStyleInfo>",
    )
    const out = patchTableStyleInfo(open, { name: 'Sales', lastColumn: true })
    expect(out.match(/<tableStyleInfo/g)).toHaveLength(1)
    expect(info(out)).toBe(
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="1" showRowStripes="1" showColumnStripes="0"/>',
    )
    expect(tableInfoOf(open, 'Sales')).toMatchObject({
      style: 'TableStyleMedium2',
      showRowStripes: true,
    })
  })

  it('puts a missing style info right after the columns, not before an extLst nested in them', () => {
    const nested = TABLE.replace(/<tableStyleInfo [^>]*\/>/, '').replace(
      '<tableColumn id="3" name="c"/>',
      '<tableColumn id="3" name="c"><extLst><ext uri="x"/></extLst></tableColumn>',
    )
    const out = patchTableStyleInfo(nested, { name: 'Sales', style: 'TableStyleLight9' })
    expect(out.indexOf('<tableStyleInfo')).toBe(
      out.indexOf('</tableColumns>') + '</tableColumns>'.length,
    )
  })

  it('finds a table by its displayName before its name', () => {
    const a = TABLE.replace(
      'name="Sales" displayName="Sales"',
      'name="Sales" displayName="Revenue"',
    )
    expect(tableKeyOf(a)).toBe('revenue')
    expect(tableInfoOf(a, 'Sales')).toBeNull()
    expect(tableInfoOf(a, 'revenue')).not.toBeNull()
    // no displayName: the name is the identity
    expect(tableKeyOf(a.replace(' displayName="Revenue"', ''))).toBe('sales')
  })
})
