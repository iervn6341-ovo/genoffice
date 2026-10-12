import { describe, expect, it } from 'vitest'

import type { WorkbookTableAdd } from '../src/shared/desktop-api'
import {
  DEFAULT_TABLE_OPTIONS,
  optionFields,
  sessionTableAt,
  tableNameProblem,
  tableOptionsOf,
} from '../src/renderer/table-design'

const entry = (extra: Partial<WorkbookTableAdd> = {}): WorkbookTableAdd => ({
  sheetId: 's1',
  area: { startRow: 1, startColumn: 1, endRow: 4, endColumn: 3 },
  name: 'Table1',
  columnNames: ['a', 'b', 'c'],
  bandedRows: true,
  ...extra,
})

describe('Table Design options', () => {
  it('reads an old entry (no options) as banded rows only', () => {
    expect(tableOptionsOf(entry())).toEqual(DEFAULT_TABLE_OPTIONS)
    expect(
      tableOptionsOf(
        entry({ bandedRows: false, bandedColumns: true, firstColumn: true, lastColumn: true }),
      ),
    ).toEqual({
      ...DEFAULT_TABLE_OPTIONS,
      bandedRows: false,
      bandedColumns: true,
      firstColumn: true,
      lastColumn: true,
    })
  })

  it('writes only the flags that are on', () => {
    expect(optionFields(DEFAULT_TABLE_OPTIONS)).toEqual({})
    expect(
      optionFields({ ...DEFAULT_TABLE_OPTIONS, firstColumn: true, bandedColumns: true }),
    ).toEqual({
      bandedColumns: true,
      firstColumn: true,
    })
  })

  it('finds the session table around a cell, on its own sheet only', () => {
    const tables = [entry(), entry({ sheetId: 's2', name: 'Table2' })]
    expect(sessionTableAt(tables, 's1', 2, 2)?.name).toBe('Table1')
    expect(sessionTableAt(tables, 's1', 1, 1)?.name).toBe('Table1')
    expect(sessionTableAt(tables, 's1', 4, 3)?.name).toBe('Table1')
    expect(sessionTableAt(tables, 's1', 5, 2)).toBeNull()
    expect(sessionTableAt(tables, 's1', 2, 4)).toBeNull()
    expect(sessionTableAt(tables, 's2', 2, 2)?.name).toBe('Table2')
  })
})

describe('Excel table name rules', () => {
  it('accepts letters, digits, underscore, period and unicode letters', () => {
    for (const name of ['Sales', '_x', 'Sales_2024', 'Q1.Totals', 'Table1', '銷售', 'A', 'AB']) {
      expect(tableNameProblem(name, []), name).toBeNull()
    }
  })

  it('refuses empty, spaced, digit-first, symbol and cell-reference names', () => {
    expect(tableNameProblem('', [])).toBe('empty')
    for (const name of [
      'my table',
      '1abc',
      'a-b',
      'a/b',
      'A1',
      'XFD1048576'.slice(0, 6),
      'R1C1',
      'R',
      'c',
      'r2c3',
    ]) {
      expect(tableNameProblem(name, []), name).toBe('invalid')
    }
    expect(tableNameProblem('x'.repeat(256), [])).toBe('invalid')
  })

  it('refuses a name another table has (any case), but lets a table keep or recase its own', () => {
    expect(tableNameProblem('sales', ['Sales', 'Other'])).toBe('taken')
    expect(tableNameProblem('Other', ['Sales', 'Other'], 'Sales')).toBe('taken')
    expect(tableNameProblem('Sales', ['Sales', 'Other'], 'Sales')).toBeNull()
    expect(tableNameProblem('SALES', ['Sales', 'Other'], 'Sales')).toBeNull()
  })
})

describe('table names Univer would refuse or rename', () => {
  it('refuses a backslash, which Excel allows but Univer rejects, and numeric or $-reference names', () => {
    for (const name of ['a\\b', '\\x', '123', '$A1', 'a:b', 'a?b', 'a*b', 'a[b]']) {
      expect(tableNameProblem(name, []), name).toBe('invalid')
    }
  })

  it("refuses a name a sheet or a defined name already has (any case), keeping the table's own", () => {
    expect(tableNameProblem('Sales', [], undefined, ['Sales', 'Budget'])).toBe('reserved')
    expect(tableNameProblem('budget', [], undefined, ['Sales', 'Budget'])).toBe('reserved')
    expect(tableNameProblem('Costs', [], undefined, ['Sales', 'Budget'])).toBeNull()
  })
})
