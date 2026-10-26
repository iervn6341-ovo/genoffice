import { describe, expect, it } from 'vitest'

import { createEditJournal, recordSheetInsert } from '../src/renderer/edit-journal'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
import { applyAiPivotAdd } from '../src/renderer/workbook-ops'

/// A sheet created in this session now takes a PivotTable (the two-phase save re-addresses it
/// by the sheet's final name), instead of the old "save first" refusal.
function fakeSheet(cells: unknown[][]) {
  const written: unknown[] = []
  const rangeOf = (row: number, column: number, height: number, width: number) => ({
    getRawValues: () =>
      cells.slice(row, row + height).map((line) => line.slice(column, column + width)),
    getNumberFormats: () =>
      cells
        .slice(row, row + height)
        .map((line) => line.slice(column, column + width).map(() => '')),
    getFormulas: () =>
      cells
        .slice(row, row + height)
        .map((line) => line.slice(column, column + width).map(() => '')),
    setValues: (matrix: unknown) => {
      written.push(matrix)
    },
  })
  return {
    written,
    getRange: (row: number, column: number, h: number, w: number) => rangeOf(row, column, h, w),
  }
}

function setup(sheetId: string) {
  const sheet = fakeSheet([
    ['Item', 'Qty'],
    ['x', 3],
    ['y', 5],
    ['x', 8],
  ])
  const journal = createEditJournal()
  recordSheetInsert(journal, 'added-1', 'Pivots')
  const runtime = {
    univerAPI: { getActiveWorkbook: () => ({ getSheetBySheetId: () => sheet }) },
  } as unknown as UniverRuntime
  const state = {
    editJournal: journal,
    file: { sheets: [], date1904: false },
    pivotDefinitions: new Map(),
  } as unknown as LazyWorkbookState
  const op = {
    op: 'add_pivot',
    sheetId,
    sourceRange: 'A1:B4',
    targetCell: 'E1',
    rowFields: ['Item'],
    values: [{ field: 'Qty', agg: 'sum' }],
  } as never
  return { runtime, state, journal, sheet, op }
}

describe('pivot on a sheet added this session', () => {
  it('is recorded against the added sheet and writes its output cells', () => {
    const { runtime, state, journal, sheet, op } = setup('added-1')
    applyAiPivotAdd(runtime, state, op)
    expect(journal.pivotAdds).toHaveLength(1)
    expect(journal.pivotAdds[0]).toMatchObject({ sheetId: 'added-1', sourceSheetId: 'added-1' })
    expect(sheet.written.length).toBeGreaterThan(0)
  })
})
