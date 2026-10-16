/**
 * Excel-style horizontal extension: columns grow on demand up to XFD without
 * touching undo history, saved content or the file's column definitions.
 */
import { describe, expect, it, vi } from 'vitest'
import { Direction, Styles, Worksheet } from '@univerjs/core'
import type { IWorksheetData } from '@univerjs/core'
import { columnIndex, columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  EXCEL_MAX_COLUMNS,
  columnCountToReach,
  ensureColumnReachable,
  extendColumns,
  planColumnExtension,
  setColumnDataSource,
} from '../src/renderer/column-extension'
import { _internals } from '../src/renderer/excel-jump-nav'
import { maxColumnIndexOfRef, resolveGoToRef, sheetNameOfRef } from '../src/renderer/goto'

function makeSheet(columnCount: number, cellData: IWorksheetData['cellData'] = {}): Worksheet {
  const snapshot: Partial<IWorksheetData> = {
    id: 'sheet-1',
    name: 'Sheet1',
    rowCount: 1000,
    columnCount,
    cellData,
    mergeData: [],
    rowData: {},
    columnData: {},
    defaultRowHeight: 19,
    defaultColumnWidth: 73,
  }
  return new Worksheet('unit-1', snapshot as IWorksheetData, new Styles())
}

/** Command service double that applies the column-count mutation like Univer. */
function fakeCommandService(worksheet: Worksheet) {
  const calls: Array<{ id: string; params: { columnCount: number } }> = []
  return {
    calls,
    service: {
      syncExecuteCommand: (id: string, params: { columnCount: number }) => {
        calls.push({ id, params })
        worksheet.setColumnCount(params.columnCount)
        return true
      },
    } as never,
  }
}

describe('column letters at the Excel boundary', () => {
  it('maps Z / AA / XFD both ways', () => {
    expect(columnLabel(25)).toBe('Z')
    expect(columnLabel(26)).toBe('AA')
    expect(columnLabel(701)).toBe('ZZ')
    expect(columnLabel(702)).toBe('AAA')
    expect(columnLabel(EXCEL_MAX_COLUMNS - 1)).toBe('XFD')
    expect(columnIndex('XFD')).toBe(EXCEL_MAX_COLUMNS - 1)
    expect(columnIndex('XFE')).toBe(EXCEL_MAX_COLUMNS)
  })
})

describe('planColumnExtension', () => {
  it('leaves a roomy grid alone', () => {
    expect(planColumnExtension(200, 10, 20)).toBeNull()
  })
  it('grows in a batch when the reach nears the last column', () => {
    expect(planColumnExtension(26, 25, 15)).toBe(78)
  })
  it('scales spare room with a wide viewport', () => {
    const narrow = planColumnExtension(100, 99, 10)!
    const wide = planColumnExtension(100, 99, 100)!
    expect(wide).toBeGreaterThan(narrow)
  })
  it('never exceeds XFD and stops once there', () => {
    expect(planColumnExtension(16300, 16299, 50)).toBe(EXCEL_MAX_COLUMNS)
    expect(planColumnExtension(EXCEL_MAX_COLUMNS, 16383, 50)).toBeNull()
  })
  it('ignores invalid reaches', () => {
    expect(planColumnExtension(26, -1)).toBeNull()
    expect(planColumnExtension(26, Number.NaN)).toBeNull()
  })
})

describe('columnCountToReach', () => {
  it('returns the exact count needed for a column and caps at XFD', () => {
    expect(columnCountToReach(26, 5)).toBeNull()
    expect(columnCountToReach(26, 26)).toBe(27)
    expect(columnCountToReach(26, 16383)).toBe(16384)
    expect(columnCountToReach(26, 99999)).toBe(16384)
  })
})

describe('extendColumns', () => {
  it('raises the column count through the mutation only', () => {
    const sheet = makeSheet(26)
    const { calls, service } = fakeCommandService(sheet)
    expect(ensureColumnReachable(service, 'unit-1', sheet, 16383)).toBe(true)
    expect(sheet.getColumnCount()).toBe(16384)
    expect(calls).toHaveLength(1)
    expect(calls[0]!.id).toBe('sheet.mutation.set-worksheet-column-count')
    // already reachable → no second mutation
    expect(ensureColumnReachable(service, 'unit-1', sheet, 100)).toBe(false)
    expect(calls).toHaveLength(1)
  })

  it('does not add cell data for the extended columns (nothing to save)', () => {
    const sheet = makeSheet(26, { 0: { 0: { v: 1 } } })
    const { service } = fakeCommandService(sheet)
    extendColumns(service, 'unit-1', sheet, 16384)
    expect(Object.keys(sheet.getConfig().cellData)).toEqual(['0'])
    expect(Object.keys(sheet.getConfig().cellData[0]!)).toEqual(['0'])
  })

  it('applies the file column definitions to the new columns only', () => {
    const sheet = makeSheet(26)
    const { service } = fakeCommandService(sheet)
    const source = vi.fn((_id: string, from: number, to: number) => {
      const data: Record<number, { w: number }> = {}
      for (let column = from; column < to; column += 1) data[column] = { w: 99 }
      return data
    })
    setColumnDataSource(source)
    try {
      extendColumns(service, 'unit-1', sheet, 30)
    } finally {
      setColumnDataSource(null)
    }
    expect(source).toHaveBeenCalledWith('sheet-1', 26, 30)
    const columnData = sheet.getConfig().columnData as Record<number, { w?: number }>
    expect(columnData[25]).toBeUndefined()
    expect(columnData[26]?.w).toBe(99)
    expect(columnData[29]?.w).toBe(99)
  })

  it('refuses to shrink or exceed the Excel limit', () => {
    const sheet = makeSheet(100)
    const { service } = fakeCommandService(sheet)
    expect(extendColumns(service, 'unit-1', sheet, 50)).toBe(false)
    extendColumns(service, 'unit-1', sheet, 99999)
    expect(sheet.getColumnCount()).toBe(EXCEL_MAX_COLUMNS)
  })
})

describe('Ctrl+Right past the materialised grid', () => {
  const { findNextGapRange } = _internals
  const start = (column: number) => ({
    startRow: 0,
    endRow: 0,
    startColumn: column,
    endColumn: column,
  })

  it('runs from an empty cell to XFD', () => {
    const sheet = makeSheet(26)
    expect(findNextGapRange(start(3), Direction.RIGHT, sheet).endColumn).toBe(16383)
  })
  it('runs from the last data cell on the grid edge to XFD', () => {
    const sheet = makeSheet(26, { 0: { 25: { v: 1 } } })
    expect(findNextGapRange(start(25), Direction.RIGHT, sheet).endColumn).toBe(16383)
  })
  it('stops at a data block that touches the grid edge', () => {
    const sheet = makeSheet(26, { 0: { 23: { v: 1 }, 24: { v: 2 }, 25: { v: 3 } } })
    expect(findNextGapRange(start(23), Direction.RIGHT, sheet).endColumn).toBe(25)
  })
  it('stays put at XFD', () => {
    const sheet = makeSheet(EXCEL_MAX_COLUMNS)
    expect(findNextGapRange(start(16383), Direction.RIGHT, sheet).endColumn).toBe(16383)
  })
})

describe('Go To far columns', () => {
  it('resolves XFD1 and reports the column to grow to', () => {
    const ref = resolveGoToRef('XFD1', [])
    expect(ref).toBe('XFD1')
    expect(maxColumnIndexOfRef(ref!)).toBe(16383)
  })
  it('rejects columns past XFD', () => {
    expect(resolveGoToRef('XFE1', [])).toBeNull()
  })
  it('reads the widest endpoint, sheet prefixes and absolute markers', () => {
    expect(maxColumnIndexOfRef("'My Sheet'!$B$2:$AZ$9")).toBe(51)
    expect(maxColumnIndexOfRef('A:XFD')).toBe(16383)
    expect(maxColumnIndexOfRef('3:9')).toBeNull()
    expect(sheetNameOfRef("'It''s'!A1")).toBe("It's")
    expect(sheetNameOfRef('A1')).toBe('')
  })
})
