import { describe, expect, it } from 'vitest'
import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { residentFormulaValues } from '../src/renderer/formula-values'

const formulaText = new Map([
  [
    'sheet-1',
    new Map([
      ['1:3', '=SUM(B1:B2)'],
      ['1:4', '=D2*2'],
      ['1:5', '=D2>0'],
      ['1:6', '=1/(B1-5)'],
      ['1:7', '=UNSUPPORTED(B1)'],
      ['1:8', '=D2*3'],
      ['1:9', '=D2+1'],
      ['1:10', '=D2+2'],
      ['1:11', '=D2+3'],
      ['1:12', '=STALE_ENGINE_ERROR(B1)'],
    ]),
  ],
  ['sheet-2', new Map([['0:1', "='Sheet 1'!E2"]])],
])

const journal: ReadonlyMap<
  string,
  ReadonlyMap<string, { row: number; column: number; formula?: string }>
> = new Map([
  [
    'sheet-1',
    new Map([
      ['1:10', { row: 1, column: 10, hasValue: false }], // formatting-only journal entry
      ['1:11', { row: 1, column: 11, formula: '=D2+3' }],
    ]),
  ],
])

const liveCells: Record<string, Record<string, CellState>> = {
  'sheet-1': {
    D2: { value: 8, rawValue: 8, formula: '=SUM(B1:B2)' },
    E2: { value: 16, rawValue: 16, formula: '=D2*2' },
    F2: { value: true, rawValue: true, formula: '=D2>0' },
    // This is a real Excel error result and should refresh its prior numeric cache.
    G2: { value: '#DIV/0!', rawValue: '#DIV/0!', formula: '=1/(B1-5)' },
    // The display still contains a usable imported cache while the model has
    // an engine failure. Saving that display as a fresh result is incorrect.
    H2: { value: 42, rawValue: '#ERROR!', formula: '=UNSUPPORTED(B1)' },
    // Formula source no longer matches the file index, so another writer owns it.
    I2: { value: 27, rawValue: 27, formula: '=D2*9' },
    // No live formula means the engine kept this as a cache-only cell.
    J2: { value: 10, rawValue: 10 },
    K2: { value: 11, rawValue: 11, formula: '=D2+2' },
    L2: { value: 12, rawValue: 12, formula: '=D2+3' },
    // Other engine errors can also be hidden by the cached-value interceptor.
    M2: { value: 99, rawValue: '#VALUE!', formula: '=STALE_ENGINE_ERROR(B1)' },
  },
  'sheet-2': {
    B1: { value: 16, rawValue: 16, formula: "='Sheet 1'!E2" },
  },
}

function reader(addresses: string[], sheetId: string): Record<string, CellState> {
  const sheet = liveCells[sheetId]
  if (!sheet) throw new Error(`unknown sheet ${sheetId}`)
  return Object.fromEntries(
    addresses.flatMap((address) => {
      const cell = sheet[address]
      return cell === undefined ? [] : [[address, cell]]
    }),
  )
}

describe('resident formula cache harvesting', () => {
  it('collects fresh number, boolean, error, chain, and cross-sheet results', () => {
    expect(residentFormulaValues(formulaText, journal, reader)).toEqual([
      { sheetId: 'sheet-1', row: 1, column: 3, value: 8 },
      { sheetId: 'sheet-1', row: 1, column: 4, value: 16 },
      { sheetId: 'sheet-1', row: 1, column: 5, value: true },
      { sheetId: 'sheet-1', row: 1, column: 6, value: { error: '#DIV/0!' } },
      { sheetId: 'sheet-1', row: 1, column: 10, value: 11 },
      { sheetId: 'sheet-2', row: 0, column: 1, value: 16 },
    ])
  })

  it('preserves a usable file cache when the live formula result is #ERROR!', () => {
    const oneFormula = new Map([['sheet-1', new Map([['1:7', '=UNSUPPORTED(B1)']])]])
    expect(residentFormulaValues(oneFormula, new Map(), reader)).toEqual([])
  })

  it('skips sheets the save transaction removed', () => {
    expect(
      residentFormulaValues(formulaText, journal, reader, (sheetId) => sheetId === 'sheet-2'),
    ).toEqual([
      { sheetId: 'sheet-1', row: 1, column: 3, value: 8 },
      { sheetId: 'sheet-1', row: 1, column: 4, value: 16 },
      { sheetId: 'sheet-1', row: 1, column: 5, value: true },
      { sheetId: 'sheet-1', row: 1, column: 6, value: { error: '#DIV/0!' } },
      { sheetId: 'sheet-1', row: 1, column: 10, value: 11 },
    ])
  })
})
