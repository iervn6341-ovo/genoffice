import type { ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  recordFileTableEdit,
  recordStructuralOp,
  toSaveTableEdits,
} from '../src/renderer/edit-journal'
import {
  designOptionsOf,
  effectiveLookOf,
  fileLookOf,
  fileTableAt,
} from '../src/renderer/file-table-design'
import { applyTableBanding } from '../src/renderer/univer-sync'
import { restyledFileTables, type LazyWorkbookState } from '../src/renderer/univer-state'

const table = (overrides: Record<string, unknown> = {}) => ({
  range: { startRow: 1, startColumn: 1, endRow: 5, endColumn: 3 },
  headerRowCount: 1,
  showRowStripes: true,
  showColumnStripes: false,
  name: 'Sales',
  styleName: 'TableStyleMedium2',
  headerFill: '#4472C4',
  stripeFill: '#D9E1F2',
  ...overrides,
})

function stateWith(tables: unknown[]): LazyWorkbookState {
  return {
    file: { sheets: [{ id: 'sheet-1', name: 'Sheet1', tables }], themeColors: null },
    editJournal: createEditJournal(),
    streamAliases: undefined,
    fileTableInfos: new Map(),
    fileTableInfoRequested: new Set(),
    fileTableShown: new Map(),
  } as unknown as LazyWorkbookState
}

describe('the file table under a cell', () => {
  const state = stateWith([
    table(),
    table({ name: undefined, range: { startRow: 8, startColumn: 1, endRow: 9, endColumn: 2 } }),
  ])

  it('is found by any cell of its range, and only by those', () => {
    expect(fileTableAt(state, 'sheet-1', 1, 1)?.name).toBe('Sales')
    expect(fileTableAt(state, 'sheet-1', 5, 3)?.name).toBe('Sales')
    expect(fileTableAt(state, 'sheet-1', 0, 1)).toBeNull()
    expect(fileTableAt(state, 'sheet-1', 3, 4)).toBeNull()
    expect(fileTableAt(state, 'other', 3, 2)).toBeNull()
  })

  it('leaves a table without a name alone (the save finds tables by name)', () => {
    expect(fileTableAt(state, 'sheet-1', 8, 1)).toBeNull()
  })

  it('leaves a table that rows or columns edits moved alone', () => {
    const moved = stateWith([table()])
    recordStructuralOp(moved.editJournal, 'sheet-1', { kind: 'insert-rows', index: 0, count: 2 })
    expect(fileTableAt(moved, 'sheet-1', 3, 2)).toBeNull()
    // an edit below the table does not move it
    const below = stateWith([table()])
    recordStructuralOp(below.editJournal, 'sheet-1', { kind: 'insert-rows', index: 20, count: 2 })
    expect(fileTableAt(below, 'sheet-1', 3, 2)?.name).toBe('Sales')
  })
})

describe('the look of a file table', () => {
  it('is what the descriptor says until the table part has been read, then what the part says', () => {
    const state = stateWith([table()])
    const ref = fileTableAt(state, 'sheet-1', 2, 2)!
    expect(fileLookOf(state, ref)).toEqual({
      style: 'TableStyleMedium2',
      bandedRows: true,
      bandedColumns: false,
      firstColumn: false,
      lastColumn: false,
    })
    state.fileTableInfos.set('sales', {
      style: 'TableStyleLight9',
      showFirstColumn: true,
      showLastColumn: false,
      showRowStripes: false,
      showColumnStripes: true,
    })
    expect(effectiveLookOf(state, ref)).toEqual({
      style: 'TableStyleLight9',
      bandedRows: false,
      bandedColumns: true,
      firstColumn: true,
      lastColumn: false,
    })
  })

  it('is the look Table Design gave it this session, when there is one', () => {
    const state = stateWith([table()])
    const ref = fileTableAt(state, 'sheet-1', 2, 2)!
    const look = {
      style: 'TableStyleDark9',
      bandedRows: true,
      bandedColumns: true,
      firstColumn: false,
      lastColumn: true,
    }
    state.fileTableShown.set(ref.key, { look, laid: [] })
    expect(effectiveLookOf(state, ref)).toEqual(look)
  })

  it('shows the options of the table: banding from the look, the structure from the table', () => {
    const state = stateWith([table({ headerRowCount: 0, totalsRowCount: 1 })])
    const ref = fileTableAt(state, 'sheet-1', 2, 2)!
    expect(designOptionsOf(ref, fileLookOf(state, ref))).toEqual({
      bandedRows: true,
      bandedColumns: false,
      firstColumn: false,
      lastColumn: false,
      headerRow: false,
      totalRow: true,
      filterButton: false,
    })
  })
})

describe('the journal keeps Table Design on a file table by name', () => {
  it('merges the newest value of a field, and saves edits of sheets that are still there', () => {
    const journal = createEditJournal()
    recordFileTableEdit(journal, { sheetId: 'sheet-1', name: 'Sales', style: 'TableStyleDark9' })
    recordFileTableEdit(journal, { sheetId: 'sheet-1', name: 'sales', bandedRows: false })
    expect(toSaveTableEdits(journal)).toEqual([
      { sheetId: 'sheet-1', name: 'sales', style: 'TableStyleDark9', bandedRows: false },
    ])
  })
})

describe('a table restyled by Table Design is not baked again by the streaming', () => {
  const baked = (): ICellData[][] => {
    const matrix: ICellData[][] = Array.from({ length: 6 }, () =>
      Array.from({ length: 4 }, () => ({})),
    )
    applyTableBanding(matrix, { startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 }, [
      table(),
    ] as never)
    return matrix
  }

  it('paints the style of the file while the descriptor is as loaded', () => {
    const matrix = baked()
    expect((matrix[1]?.[1]?.s as { bg?: { rgb?: string } }).bg?.rgb).toBe('#4472C4')
  })

  it('paints nothing once the descriptor is marked', () => {
    const descriptor = table()
    restyledFileTables.add(descriptor)
    const matrix: ICellData[][] = Array.from({ length: 6 }, () =>
      Array.from({ length: 4 }, () => ({})),
    )
    applyTableBanding(matrix, { startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 }, [
      descriptor,
    ] as never)
    expect(matrix.flat().every((cell) => cell.s === undefined)).toBe(true)
  })
})
