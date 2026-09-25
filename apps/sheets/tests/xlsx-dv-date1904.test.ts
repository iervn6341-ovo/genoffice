/**
 * C2: a Date validation authored in the panel stores ISO strings; the save
 * turned them into 1900-system serials even when the workbook counts from
 * 1904 (workbookPr/@date1904), so Excel read 2026 bounds as 2030/2031.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { applyDvRules } from '@genoffice/xlsx-gateway/gateway/xlsx-dv'
import {
  applyCellEditsToXlsx,
  workbookUsesDate1904,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

const SHEET = '<worksheet><sheetData/></worksheet>'
const range = { startRow: 0, endRow: 4, startColumn: 1, endColumn: 1 }
const dateRule = {
  ranges: [range],
  rule: { type: 'date', operator: 'between', formula1: '2026-01-01', formula2: '2026-12-31' },
}
// Excel serials for 2026-01-01 / 2026-12-31
const S1900 = [46023, 46387] as const

describe('date validation epoch (C2)', () => {
  it('1900 system keeps the 1900 serials', () => {
    const xml = applyDvRules(SHEET, [dateRule])
    expect(xml).toContain(`<formula1>${S1900[0]}</formula1>`)
    expect(xml).toContain(`<formula2>${S1900[1]}</formula2>`)
  })

  it('1904 system subtracts 1462 days from panel date strings', () => {
    const xml = applyDvRules(SHEET, [dateRule], { date1904: true })
    expect(xml).toContain(`<formula1>${S1900[0] - 1462}</formula1>`)
    expect(xml).toContain(`<formula2>${S1900[1] - 1462}</formula2>`)
  })

  it('1904 system keeps numeric (already file-epoch) bounds and date-times', () => {
    const xml = applyDvRules(
      SHEET,
      [
        { ranges: [range], rule: { type: 'date', operator: 'greaterThan', formula1: '44561' } },
        {
          ranges: [{ ...range, startColumn: 2, endColumn: 2 }],
          rule: { type: 'date', operator: 'lessThan', formula1: '2026-01-01 12:00' },
        },
        {
          ranges: [{ ...range, startColumn: 3, endColumn: 3 }],
          rule: { type: 'time', operator: 'lessThan', formula1: '06:00' },
        },
      ],
      { date1904: true },
    )
    expect(xml).toContain('<formula1>44561</formula1>')
    expect(xml).toContain(`<formula1>${S1900[0] - 1462 + 0.5}</formula1>`)
    expect(xml).toContain('<formula1>0.25</formula1>')
  })

  it('append mode honors the epoch too', () => {
    const xml = applyDvRules(SHEET, [dateRule], { append: true, date1904: true })
    expect(xml).toContain(`<formula1>${S1900[0] - 1462}</formula1>`)
  })

  it('reads workbookPr/@date1904 in its ST_OnOff spellings', () => {
    expect(workbookUsesDate1904('<workbook><workbookPr date1904="1"/></workbook>')).toBe(true)
    expect(workbookUsesDate1904('<x:workbook><x:workbookPr date1904="true"/></x:workbook>')).toBe(
      true,
    )
    expect(workbookUsesDate1904('<workbook><workbookPr date1904="0"/></workbook>')).toBe(false)
    expect(workbookUsesDate1904('<workbook><workbookPr/></workbook>')).toBe(false)
  })

  async function saveWithDateRule(date1904: boolean): Promise<string> {
    const zip = await JSZip.loadAsync(await buildEditFixture())
    if (date1904) {
      const wb = (await zip.file('xl/workbook.xml')!.async('text')).replace(
        '<sheets>',
        '<workbookPr date1904="1"/><sheets>',
      )
      zip.file('xl/workbook.xml', wb)
    }
    const source = await zip.generateAsync({ type: 'nodebuffer' })
    const mutation = await applyCellEditsToXlsx(
      source,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [{ sheetName: 'Data', rules: [dateRule] }],
    )
    const out = await JSZip.loadAsync(mutation.buffer)
    return (await out.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
  }

  it('round-trip: the save path reads the epoch from the workbook', async () => {
    expect(await saveWithDateRule(false)).toContain(`<formula1>${S1900[0]}</formula1>`)
    const xml1904 = await saveWithDateRule(true)
    expect(xml1904).toContain(`<formula1>${S1900[0] - 1462}</formula1>`)
    expect(xml1904).toContain(`<formula2>${S1900[1] - 1462}</formula2>`)
  })
})
