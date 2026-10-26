import { describe, expect, it } from 'vitest'

import { addedSheetFinalNames, remapHeldToSavedSheets } from '../src/renderer/held-sheet-remap'

describe('addedSheetFinalNames', () => {
  it('maps added and duplicated sheets to their final names (renames win)', () => {
    const names = addedSheetFinalNames([
      { kind: 'add-sheet', sheetId: 'a', name: 'Sheet2' },
      { kind: 'duplicate-sheet', sheetId: 'b', name: 'Copy' },
      { kind: 'rename-sheet', sheetId: 'a', newName: 'Pivots' },
      { kind: 'rename-sheet', sheetId: 'existing', newName: 'Ignored' },
      { kind: 'remove-sheet', sheetId: 'c' },
    ])
    expect([...names]).toEqual([
      ['a', 'Pivots'],
      ['b', 'Copy'],
    ])
  })
})

describe('remapHeldToSavedSheets', () => {
  const saved = [
    { id: 'sheet-1', name: 'Data' },
    { id: 'sheet-7', name: 'Pivots' },
  ]
  const names = new Map([['tmp', 'Pivots']])

  it('re-addresses items on sheets created in this save and leaves existing ids alone', () => {
    const out = remapHeldToSavedSheets(
      [
        { sheetId: 'tmp', sourceSheetId: 'sheet-1', tag: 'p' },
        { sheetId: 'sheet-1', sourceSheetId: 'sheet-1', tag: 'q' },
      ],
      names,
      saved,
    )
    expect(out).toEqual([
      { sheetId: 'sheet-7', sourceSheetId: 'sheet-1', tag: 'p' },
      { sheetId: 'sheet-1', sourceSheetId: 'sheet-1', tag: 'q' },
    ])
  })

  it('also remaps a source sheet that was created in this save, and plain tables without one', () => {
    const out = remapHeldToSavedSheets([{ sheetId: 'sheet-1', sourceSheetId: 'tmp' }], names, saved)
    expect(out[0]).toEqual({ sheetId: 'sheet-1', sourceSheetId: 'sheet-7' })
    const tables = remapHeldToSavedSheets([{ sheetId: 'tmp' }], names, saved)
    expect(tables).toEqual([{ sheetId: 'sheet-7' }])
    expect('sourceSheetId' in tables[0]!).toBe(false)
  })

  it('matches names the way Excel does (case-insensitively)', () => {
    const out = remapHeldToSavedSheets([{ sheetId: 'tmp' }], new Map([['tmp', 'pivots']]), saved)
    expect(out[0]!.sheetId).toBe('sheet-7')
  })

  it('throws instead of guessing when the saved file lacks the sheet', () => {
    expect(() => remapHeldToSavedSheets([{ sheetId: 'tmp' }], names, [saved[0]!])).toThrow(/Pivots/)
  })

  it('does not mutate its input', () => {
    const input = [{ sheetId: 'tmp' }]
    remapHeldToSavedSheets(input, names, saved)
    expect(input[0]!.sheetId).toBe('tmp')
  })
})
