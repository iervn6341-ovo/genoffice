/**
 * C3: formulas typed in the grid were saved as a bare <f> — only MCP-verified
 * formulas got a cached <v> — so readers without a formula engine (Quick Look,
 * pandas, mobile viewers) showed empty cells until Excel recalculated.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { journaledFormulaValues } from '../src/renderer/formula-values'
import { buildStructureFixture } from './fixture-builder'

const reader =
  (cells: Record<string, Record<string, CellState>>) => (addresses: string[], sheetId: string) => {
    const sheet = cells[sheetId]
    if (!sheet) throw new Error(`no sheet ${sheetId}`)
    return Object.fromEntries(addresses.filter((a) => a in sheet).map((a) => [a, sheet[a]!]))
  }

const journal = (entries: Record<string, { row: number; column: number; formula?: string }[]>) =>
  new Map(
    Object.entries(entries).map(([sheetId, cells]) => [
      sheetId,
      new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell])),
    ]),
  )

describe('journaledFormulaValues (C3)', () => {
  it('caches number, string, boolean and error results of typed formulas', () => {
    const read = reader({
      s1: {
        B1: { value: 14, formula: '=A1*2' },
        B2: { value: 'hi', formula: '="h"&"i"' },
        B3: { value: true, formula: '=A1>0' },
        B4: { value: '#DIV/0!', formula: '=1/0' },
        AA10: { value: 3, formula: '=1+2' },
      },
    })
    const values = journaledFormulaValues(
      journal({
        s1: [
          { row: 0, column: 1, formula: '=A1*2' },
          { row: 1, column: 1, formula: '="h"&"i"' },
          { row: 2, column: 1, formula: '=A1>0' },
          { row: 3, column: 1, formula: '=1/0' },
          { row: 9, column: 26, formula: '=1+2' },
          { row: 5, column: 0 }, // plain value edit: not a formula
        ],
      }),
      read,
    )
    expect(values).toEqual([
      { sheetId: 's1', row: 0, column: 1, value: 14 },
      { sheetId: 's1', row: 1, column: 1, value: 'hi' },
      { sheetId: 's1', row: 2, column: 1, value: true },
      { sheetId: 's1', row: 3, column: 1, value: { error: '#DIV/0!' } },
      { sheetId: 's1', row: 9, column: 26, value: 3 },
    ])
  })

  it('skips replaced formulas, unsettled cells, engine failures and removed sheets', () => {
    const read = reader({
      s1: {
        B1: { value: 7, formula: '=A1*3' },
        B2: { value: null, formula: '=A1' },
        B3: { value: '#ERROR!', formula: '=NOPE()' },
      },
    })
    const values = journaledFormulaValues(
      journal({
        s1: [
          { row: 0, column: 1, formula: '=A1*2' },
          { row: 1, column: 1, formula: '=A1' },
          { row: 2, column: 1, formula: '=NOPE()' },
        ],
        gone: [{ row: 0, column: 0, formula: '=1' }],
        removed: [{ row: 0, column: 0, formula: '=1' }],
      }),
      read,
      (sheetId) => sheetId === 'removed',
    )
    expect(values).toEqual([])
  })

  it('honors the per-save cell budget', () => {
    const read = reader({
      s1: { A1: { value: 1, formula: '=1' }, A2: { value: 2, formula: '=2' } },
    })
    const values = journaledFormulaValues(
      journal({
        s1: [
          { row: 0, column: 0, formula: '=1' },
          { row: 1, column: 0, formula: '=2' },
        ],
      }),
      read,
      undefined,
      1,
    )
    expect(values).toHaveLength(1)
  })
})

describe('cached <v> XML per result type (C3)', () => {
  async function savedCell(value: string | number | boolean | { error: string }) {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      [{ sheetName: 'Data', row: 0, column: 0, writeValue: true, cell: { value: 1 } }],
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
      [{ sheetName: 'Data', cells: [{ row: 1, column: 3, value }] }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const xml = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    return /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
  }

  it('number: no t attribute', async () => {
    const cell = await savedCell(42)
    expect(cell).toContain('<v>42</v>')
    expect(cell).not.toMatch(/\st="/)
    expect(cell).toContain('<f>')
  })
  it('string: t="str"', async () => {
    const cell = await savedCell('abc')
    expect(cell).toContain('t="str"')
    expect(cell).toContain('<v>abc</v>')
  })
  it('boolean: t="b" with 1/0', async () => {
    expect(await savedCell(true)).toMatch(/t="b"[^>]*>[\s\S]*<v>1<\/v>/)
    expect(await savedCell(false)).toContain('<v>0</v>')
  })
  it('error: t="e"', async () => {
    const cell = await savedCell({ error: '#N/A' })
    expect(cell).toContain('t="e"')
    expect(cell).toContain('<v>#N/A</v>')
  })

  it('a formula typed into a new cell gets <f> and its cached <v> in one save', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildStructureFixture(),
      [
        {
          sheetName: 'Data',
          row: 9,
          column: 5,
          writeValue: true,
          cell: { value: null, formula: '=1+2' },
        },
      ],
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
      [{ sheetName: 'Data', cells: [{ row: 9, column: 5, value: 3 }] }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const xml = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    const cell = /<c[^>]*r="F10"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('<f>1+2</f>')
    expect(cell).toContain('<v>3</v>')
  })
})
