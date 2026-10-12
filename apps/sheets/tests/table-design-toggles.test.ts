import { describe, expect, it } from 'vitest'

import type { WorkbookTableAdd } from '../src/shared/desktop-api'
import {
  checkResize,
  dataRowsOf,
  DEFAULT_TABLE_OPTIONS,
  optionFields,
  parseAreaText,
  planTableChange,
  tableOptionsOf,
  totalsFormula,
  totalsKindOf,
  univerTableRange,
} from '../src/renderer/table-design'

const entry = (extra: Partial<WorkbookTableAdd> = {}): WorkbookTableAdd => ({
  sheetId: 's1',
  area: { startRow: 1, startColumn: 1, endRow: 4, endColumn: 3 },
  name: 'Table1',
  columnNames: ['a', 'b', 'c'],
  bandedRows: true,
  ...extra,
})

describe('Header Row, Total Row and Filter Button', () => {
  it('read an old entry as Excel’s default and write only what differs from it', () => {
    expect(tableOptionsOf(entry())).toMatchObject({
      headerRow: true,
      totalRow: false,
      filterButton: true,
    })
    const options = tableOptionsOf(
      entry({ headerRow: false, totalsRow: true, filterButton: false }),
    )
    expect(options).toMatchObject({ headerRow: false, totalRow: true, filterButton: false })
    expect(optionFields(options)).toEqual({
      headerRow: false,
      totalsRow: true,
      filterButton: false,
    })
    expect(optionFields(DEFAULT_TABLE_OPTIONS)).toEqual({})
  })

  it('maps the journal area (the table as the file holds it) to the range Univer draws', () => {
    const area = { startRow: 2, startColumn: 1, endRow: 8, endColumn: 3 }
    expect(univerTableRange(area, { headerRow: true, totalRow: false })).toEqual(area)
    // a totals row stays outside Univer's table, so Table1[Col] does not include the cell that sums it
    expect(univerTableRange(area, { headerRow: true, totalRow: true }).endRow).toBe(7)
    // without a header row Univer's table reaches up over the empty row the header left
    expect(univerTableRange(area, { headerRow: false, totalRow: false }).startRow).toBe(1)
    expect(dataRowsOf(area, { headerRow: true, totalRow: true })).toEqual({ first: 3, last: 7 })
    expect(dataRowsOf(area, { headerRow: false, totalRow: false })).toEqual({ first: 2, last: 8 })
  })

  it('writes Excel’s SUBTOTAL for a totals cell: a sum of numbers, a count of anything else', () => {
    expect(totalsKindOf([1, 2, 3])).toBe('sum')
    expect(totalsKindOf([1, '', null])).toBe('sum')
    expect(totalsKindOf(['x', 'y'])).toBe('count')
    expect(totalsKindOf([1, 'x'])).toBe('count')
    expect(totalsKindOf([])).toBe('count')
    // Excel 16 (COM): =SUBTOTAL(109,[Price]) for a number column, =SUBTOTAL(103,[Txt]) for text
    expect(totalsFormula('sum', 'Table1', 'Price')).toBe('=SUBTOTAL(109,Table1[Price])')
    expect(totalsFormula('count', 'T_2', 'Unit Price')).toBe('=SUBTOTAL(103,T_2[Unit Price])')
    // a special character in a column name is escaped with a quote
    expect(totalsFormula('sum', 'T', '#1')).toBe("=SUBTOTAL(109,T['#1])")
  })
})

describe('what turning Header Row and Total Row on and off does to the cells (read from Excel 16)', () => {
  // A table at A2:C4: header row 1, data rows 2–3 (0-based)
  const area = { startRow: 1, startColumn: 0, endRow: 3, endColumn: 2 }
  const cells = (grid: Record<string, unknown>) => (row: number, column: number) =>
    grid[`${row}:${column}`] ?? null
  const base = {
    area,
    columnNames: ['Item', 'Qty', 'Price'],
    tableName: 'Table1',
    totalLabel: 'Total',
    maxRows: 1000,
  }
  const grid = {
    '1:0': 'Item',
    '1:1': 'Qty',
    '1:2': 'Price',
    '2:0': 'a',
    '2:1': 1,
    '2:2': 10,
    '3:0': 'b',
    '3:1': 2,
    '3:2': 20,
  }
  const on = { headerRow: true, totalRow: true }
  const off = { headerRow: true, totalRow: false }

  it('Total Row on: "Total" under the first column and a SUBTOTAL under the last, one row lower', () => {
    const change = planTableChange({ ...base, from: off, to: on, cellAt: cells(grid) })
    if ('error' in change) throw new Error(change.error)
    expect(change.area).toEqual({ ...area, endRow: 4 })
    expect(change.writes).toEqual([
      { row: 4, column: 0, value: 'Total' },
      { row: 4, column: 2, value: '=SUBTOTAL(109,Table1[Price])' },
    ])
    expect(change.totals).toEqual([{ label: 'Total' }, null, { fn: 'sum' }])
  })

  it('Total Row over a text column counts, like Excel (SUBTOTAL 103)', () => {
    const change = planTableChange({
      ...base,
      from: off,
      to: on,
      cellAt: cells({ ...grid, '2:2': 'x', '3:2': 'y' }),
    })
    if ('error' in change) throw new Error(change.error)
    expect(change.writes[1]).toEqual({ row: 4, column: 2, value: '=SUBTOTAL(103,Table1[Price])' })
    expect(change.totals?.[2]).toEqual({ fn: 'count' })
  })

  it('refuses Total Row when the row under the table is not empty (Excel takes the rows in)', () => {
    expect(
      planTableChange({ ...base, from: off, to: on, cellAt: cells({ ...grid, '4:1': 'below' }) }),
    ).toEqual({ error: 'noRoomBelow' })
    expect(
      planTableChange({ ...base, maxRows: 4, from: off, to: on, cellAt: cells(grid) }),
    ).toEqual({ error: 'noRoomBelow' })
  })

  it('a one-column table gets only the SUBTOTAL', () => {
    const change = planTableChange({
      ...base,
      area: { ...area, endColumn: 0 },
      columnNames: ['Item'],
      from: off,
      to: on,
      cellAt: cells(grid),
    })
    if ('error' in change) throw new Error(change.error)
    expect(change.writes).toEqual([{ row: 4, column: 0, value: '=SUBTOTAL(103,Table1[Item])' }])
  })

  it('Total Row off empties the row and the table is one row shorter', () => {
    const change = planTableChange({
      ...base,
      area: { ...area, endRow: 4 },
      from: on,
      to: off,
      cellAt: cells({ ...grid, '4:0': 'Total', '4:2': '=SUBTOTAL(109,Table1[Price])' }),
    })
    if ('error' in change) throw new Error(change.error)
    expect(change.area).toEqual(area)
    expect(change.writes).toEqual([
      { row: 4, column: 0, value: null },
      { row: 4, column: 1, value: null },
      { row: 4, column: 2, value: null },
    ])
  })

  it('Header Row off empties the header cells; the table starts one row lower and keeps its names', () => {
    const change = planTableChange({
      ...base,
      from: off,
      to: { headerRow: false, totalRow: false },
      cellAt: cells({ ...grid, '1:1': 'Quantity' }),
    })
    if ('error' in change) throw new Error(change.error)
    expect(change.area).toEqual({ ...area, startRow: 2 })
    // a header edited since the table was made is the name now
    expect(change.columnNames).toEqual(['Item', 'Quantity', 'Price'])
    expect(change.writes.map((write) => write.value)).toEqual([null, null, null])
  })

  it('Header Row back on: the names go into the row above (which must be empty); otherwise refused', () => {
    const headerless = { ...area, startRow: 2 }
    const none = { headerRow: false, totalRow: false }
    const change = planTableChange({
      ...base,
      area: headerless,
      from: none,
      to: off,
      cellAt: cells({ '2:0': 'a', '3:0': 'b' }),
    })
    if ('error' in change) throw new Error(change.error)
    expect(change.area).toEqual(area)
    expect(change.writes).toEqual([
      { row: 1, column: 0, value: 'Item' },
      { row: 1, column: 1, value: 'Qty' },
      { row: 1, column: 2, value: 'Price' },
    ])
    expect(
      planTableChange({
        ...base,
        area: headerless,
        from: none,
        to: off,
        cellAt: cells({ '1:1': 'x' }),
      }),
    ).toEqual({ error: 'noRoomAbove' })
    expect(
      planTableChange({
        ...base,
        area: { ...area, startRow: 0 },
        from: none,
        to: off,
        cellAt: cells({}),
      }),
    ).toEqual({ error: 'noRoomAbove' })
  })

  it('needs at least one row of data when the header goes', () => {
    expect(
      planTableChange({
        ...base,
        area: { ...area, endRow: 1 },
        from: off,
        to: { headerRow: false, totalRow: false },
        cellAt: cells(grid),
      }),
    ).toEqual({ error: 'needsData' })
  })
})

describe('Header Row off names the columns from their header cells', () => {
  const area = { startRow: 1, startColumn: 0, endRow: 3, endColumn: 2 }
  const off = { headerRow: false, totalRow: false }
  const on = { headerRow: true, totalRow: false }
  const run = (headers: Record<number, unknown>) =>
    planTableChange({
      area,
      columnNames: ['Item', 'Qty', 'Price'],
      from: on,
      to: off,
      tableName: 'Table1',
      totalLabel: 'Total',
      maxRows: 1000,
      cellAt: (row, column) => (row === 1 ? (headers[column] ?? null) : 1),
    })

  it('makes headers that were edited into the same word unique, in any case, the way Excel names columns', () => {
    const change = run({ 0: 'Same', 1: 'same', 2: 'SAME' })
    if ('error' in change) throw new Error(change.error)
    expect(change.columnNames).toEqual(['Same', 'same2', 'SAME3'])
  })

  it('keeps a column name at most 255 characters long', () => {
    const change = run({ 0: 'x'.repeat(400), 1: 'Qty', 2: 'Price' })
    if ('error' in change) throw new Error(change.error)
    expect(change.columnNames[0]!.length).toBeLessThanOrEqual(255)
  })
})

describe('Resize Table', () => {
  it('reads a typed range, with or without dollar signs and spaces', () => {
    expect(parseAreaText('A1:D10')).toEqual({
      startRow: 0,
      endRow: 9,
      startColumn: 0,
      endColumn: 3,
    })
    expect(parseAreaText(' $b$2 : $C$5 ')).toEqual({
      startRow: 1,
      endRow: 4,
      startColumn: 1,
      endColumn: 2,
    })
    expect(parseAreaText('D10:A1')).toEqual({
      startRow: 0,
      endRow: 9,
      startColumn: 0,
      endColumn: 3,
    })
    // what the Resize Table box holds is a reference: it starts with an = as often as not
    expect(parseAreaText('=$A$1:$C$5')).toEqual({
      startRow: 0,
      endRow: 4,
      startColumn: 0,
      endColumn: 2,
    })
    expect(parseAreaText('Sheet1!A1:B2')).toBeNull()
    expect(parseAreaText('')).toBeNull()
    expect(parseAreaText('A0:B2')).toBeNull()
  })

  it('keeps the top-left cell, needs a row of data, and waits for the Total Row to go', () => {
    const area = { startRow: 1, startColumn: 1, endRow: 4, endColumn: 3 }
    const options = { headerRow: true, totalRow: false }
    expect(checkResize(area, { ...area, endRow: 9, endColumn: 5 }, options)).toBeNull()
    expect(checkResize(area, { ...area, endRow: 2 }, options)).toBeNull()
    expect(checkResize(area, { ...area, endRow: 1 }, options)).toBe('needsData')
    expect(checkResize(area, { ...area, startRow: 0 }, options)).toBe('resizeStart')
    expect(checkResize(area, { ...area, startColumn: 0 }, options)).toBe('resizeStart')
    expect(checkResize(area, area, { headerRow: true, totalRow: true })).toBe('resizeTotals')
    // a width nobody means (a typo such as A1:ZZ9) is refused rather than turned into a table
    expect(checkResize(area, { ...area, endColumn: 1200 }, options)).toBe('resizeOutside')
    // a headerless table: one data row is enough
    expect(
      checkResize(area, { ...area, endRow: 1 }, { headerRow: false, totalRow: false }),
    ).toBeNull()
  })
})
