import { CellValueType } from '@univerjs/core'
import JSZip from 'jszip'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseTextToColumnsAddress,
  planTextToColumns,
  splitDelimitedCell,
  splitFixedWidthCell,
  textToColumnsAddress,
  parseTextToColumnsDate,
  parseTextToColumnsNumber,
  textToColumnsCell,
  type TextToColumnsOptions,
  type TextToColumnsSeed,
} from '../src/renderer/text-to-columns'

const comma: TextToColumnsOptions = {
  tab: false,
  semicolon: false,
  comma: true,
  space: false,
  other: '',
  collapse: false,
  qualifier: 'double',
}

const seed: TextToColumnsSeed = {
  workbookId: 'book',
  sheetId: 'sheet',
  sourceColumn: 0,
  rows: [
    { row: 4, text: '"A, B",003,4' },
    { row: 5, text: '"line\nbreak",5,=1+2' },
  ],
  maxRows: 100,
  maxColumns: 20,
}

describe('Text to Columns planning', () => {
  it('parses Excel qualifiers, escaped quotes, empty fields and embedded newlines', () => {
    expect(splitDelimitedCell('"A, B","a""b",,end,', comma)).toEqual({
      fields: ['A, B', 'a"b', '', 'end', ''],
    })
    expect(splitDelimitedCell('"unclosed,3', comma)).toEqual({ error: 'qualifier' })
    expect(splitDelimitedCell('a,b', { ...comma, comma: false })).toEqual({ error: 'delimiter' })
    expect(splitDelimitedCell('a,b', { ...comma, other: 'ab' })).toEqual({ error: 'delimiter' })
    expect(splitDelimitedCell('a,, b', { ...comma, collapse: true, space: true })).toEqual({
      fields: ['a', 'b'],
    })
    expect(splitDelimitedCell('"A, B",3', { ...comma, qualifier: 'none' })).toEqual({
      fields: ['"A', ' B"', '3'],
    })
  })

  it('uses one fixed-width boundary set for unequal rows, including empty trailing fields', () => {
    expect(splitFixedWidthCell('ABC0036', [3, 6])).toEqual({ fields: ['ABC', '003', '6'] })
    expect(splitFixedWidthCell('XY', [3, 6])).toEqual({ fields: ['XY', '', ''] })
    expect(splitFixedWidthCell('😀B0036', [2, 5])).toEqual({ fields: ['😀B', '003', '6'] })
    expect(splitFixedWidthCell('A', [])).toEqual({ error: 'breaks' })
    expect(splitFixedWidthCell('A', [4, 4])).toEqual({ error: 'breaks' })
    expect(splitFixedWidthCell('A', [2, 1])).toEqual({ error: 'breaks' })

    const planned = planTextToColumns(
      {
        ...seed,
        rows: [
          { row: 4, text: 'ABC0036' },
          { row: 5, text: 'XY' },
        ],
      },
      { ...comma, fixedWidth: { breaks: [3, 6] } },
      { row: 4, column: 0 },
      ['general', 'text', 'general'],
      () => 'free',
    )
    expect(planned.plan?.writes.map((entry) => entry.values.map((row) => row[0]?.v))).toEqual([
      ['ABC', 'XY'],
      ['003', ''],
      [6, ''],
    ])
    expect(
      planTextToColumns(
        { ...seed, rows: [{ row: 4, text: 'ABC0036' }] },
        { ...comma, fixedWidth: { breaks: [3, 6] } },
        { row: 4, column: 0 },
        [],
        (_, column) => (column === 1 ? 'blocked' : 'free'),
      ).error,
    ).toBe('occupied')
  })

  it('plans only safe destination writes, preserving skipped columns and formula-looking text', () => {
    const planned = planTextToColumns(
      seed,
      comma,
      { row: 4, column: 0 },
      ['text', 'general', 'text'],
      () => 'free',
    )
    expect(planned.error).toBeUndefined()
    expect(planned.plan?.writes).toHaveLength(3)
    expect(planned.plan?.writes[0]?.values[0]?.[0]).toMatchObject({
      v: 'A, B',
      t: CellValueType.FORCE_STRING,
      f: null,
    })
    expect(planned.plan?.writes[1]?.values).toEqual([
      [expect.objectContaining({ v: 3, t: CellValueType.NUMBER })],
      [expect.objectContaining({ v: 5, t: CellValueType.NUMBER })],
    ])
    expect(planned.plan?.writes[2]?.values[1]?.[0]).toMatchObject({
      v: '=1+2',
      t: CellValueType.FORCE_STRING,
      f: null,
    })
    expect(
      planTextToColumns(
        seed,
        comma,
        { row: 4, column: 0 },
        ['general', 'skip', 'text'],
        () => 'free',
      ).plan?.writes.map((entry) => entry.column),
    ).toEqual([0, 2])
    expect(
      planTextToColumns(seed, comma, { row: 4, column: 0 }, ['skip', 'skip', 'skip'], () => 'free')
        .error,
    ).toBe('source')
  })

  it('rejects occupied output and overlaps before any mutation', () => {
    expect(
      planTextToColumns(seed, comma, { row: 4, column: 0 }, [], (_, column) =>
        column === 1 ? 'blocked' : 'free',
      ).error,
    ).toBe('occupied')
    expect(
      planTextToColumns(seed, comma, { row: 4, column: 1 }, [], () => 'free').plan,
    ).toBeDefined()
    expect(planTextToColumns(seed, comma, { row: 5, column: 0 }, [], () => 'free').error).toBe(
      'destination',
    )
    expect(planTextToColumns(seed, comma, { row: 4, column: 19 }, [], () => 'free').error).toBe(
      'destination',
    )
  })

  it('uses strict A1 addresses and General numeric conversion', () => {
    expect(parseTextToColumnsAddress('$AA$51')).toEqual({ row: 50, column: 26 })
    expect(textToColumnsAddress(50, 26)).toBe('AA51')
    expect(parseTextToColumnsAddress('Sheet1!A1')).toBeNull()
    expect(parseTextToColumnsAddress('A0')).toBeNull()
    expect(textToColumnsCell('001.25', 'general')).toMatchObject({
      v: 1.25,
      t: CellValueType.NUMBER,
    })
    expect(textToColumnsCell('12/13', 'general')).toMatchObject({
      v: '12/13',
      t: CellValueType.STRING,
    })
  })

  it('matches the Excel Mac quoted-CSV specimen while protecting its destination formula', async () => {
    const directory = join(
      fileURLToPath(new URL('../../..', import.meta.url)),
      'docs/qa/office-mac-field-evidence-2026-10-13/sheets-text-to-columns',
    )
    const before = await JSZip.loadAsync(
      await readFile(join(directory, 'excel-mac-text-to-columns-before.xlsx')),
    )
    const after = await JSZip.loadAsync(
      await readFile(join(directory, 'excel-mac-text-to-columns-after.xlsx')),
    )
    const beforeSheet = await before.file('xl/worksheets/sheet1.xml')!.async('string')
    const afterSheet = await after.file('xl/worksheets/sheet1.xml')!.async('string')
    const beforeStrings = await before.file('xl/sharedStrings.xml')!.async('string')
    const afterStrings = await after.file('xl/sharedStrings.xml')!.async('string')
    expect(beforeSheet).toContain('<f>10+5</f><v>15</v>')
    expect(beforeStrings).toContain('"A, B",003,4')
    expect(beforeStrings).toContain('"C,D",005,6')
    expect(afterStrings).toContain('A, B')
    expect(afterStrings).toContain('003')
    expect(afterSheet).not.toContain('<f>10+5</f>')

    const nativeSeed: TextToColumnsSeed = {
      ...seed,
      rows: [
        { row: 0, text: '"A, B",003,4' },
        { row: 1, text: '"C,D",005,6' },
      ],
    }
    const overwritePlan = planTextToColumns(
      nativeSeed,
      comma,
      { row: 0, column: 0 },
      ['general', 'text', 'general'],
      (row, column) => (row === 0 && column === 1 ? 'overwrite' : 'free'),
    )
    // Excel prompts and then replaces the B1 formula: the plan records it for confirmation
    expect(overwritePlan.error).toBeUndefined()
    expect(overwritePlan.plan?.overwrites).toEqual([{ row: 0, column: 1 }])
    const safe = planTextToColumns(
      nativeSeed,
      comma,
      { row: 0, column: 3 },
      ['general', 'text', 'general'],
      () => 'free',
    )
    expect(safe.plan?.writes.map((write) => write.values.map((row) => row[0]?.v))).toEqual([
      ['A, B', 'C,D'],
      ['003', '005'],
      [4, 6],
    ])
  })

  it('parses Date columns in every Y/M/D order, two-digit years and month names', () => {
    // 2024-12-13 is serial 45639 in the 1900 system
    expect(parseTextToColumnsDate('12/13/2024', 'mdy')).toBe(45639)
    expect(parseTextToColumnsDate('13.12.2024', 'dmy')).toBe(45639)
    expect(parseTextToColumnsDate('2024-12-13', 'ymd')).toBe(45639)
    expect(parseTextToColumnsDate('20241213', 'ymd')).toBe(45639)
    expect(parseTextToColumnsDate('12 2024 13', 'myd')).toBe(45639)
    expect(parseTextToColumnsDate('13 12 2024', 'dmy')).toBe(45639)
    expect(parseTextToColumnsDate('13 2024 12', 'dym')).toBe(45639)
    expect(parseTextToColumnsDate('2024 13 12', 'ydm')).toBe(45639)
    expect(parseTextToColumnsDate('13-Dec-2024', 'dmy')).toBe(45639)
    expect(parseTextToColumnsDate('12/13/24', 'mdy')).toBe(45639)
    expect(parseTextToColumnsDate('12/13/99', 'mdy')).toBe(36507)
    expect(parseTextToColumnsDate('01/01/1900', 'mdy')).toBe(1)
    expect(parseTextToColumnsDate('03/01/1900', 'mdy')).toBe(61)
    expect(parseTextToColumnsDate('12/13/2024', 'mdy', true)).toBe(45639 - 1462)
    expect(parseTextToColumnsDate('13/12/2024', 'mdy')).toBeNull()
    expect(parseTextToColumnsDate('02/30/2024', 'mdy')).toBeNull()
    expect(parseTextToColumnsDate('hello', 'mdy')).toBeNull()
  })

  it('writes Date cells with the system short-date pattern and leaves bad dates as text', () => {
    expect(textToColumnsCell('12/13/2024', 'date-mdy')).toMatchObject({
      v: 45639,
      t: CellValueType.NUMBER,
      s: { n: { pattern: 'm/d/yyyy' } },
    })
    expect(textToColumnsCell('13/13/2024', 'date-mdy')).toMatchObject({
      v: '13/13/2024',
      t: CellValueType.STRING,
    })
    expect(textToColumnsCell('12/13/2024', 'text')).toMatchObject({ t: CellValueType.FORCE_STRING })
  })

  it('applies Advanced decimal / thousands separators and trailing minus', () => {
    expect(parseTextToColumnsNumber('1,234.5')).toBe(1234.5)
    expect(parseTextToColumnsNumber('1,23')).toBeNull()
    expect(parseTextToColumnsNumber('12,345,678')).toBe(12345678)
    const european = { decimal: ',', thousands: '.', trailingMinus: false }
    expect(parseTextToColumnsNumber('1.234,5', european)).toBe(1234.5)
    expect(parseTextToColumnsNumber('3,14', european)).toBe(3.14)
    expect(parseTextToColumnsNumber('1,234.5', european)).toBeNull()
    expect(parseTextToColumnsNumber('5-', { ...european, trailingMinus: true })).toBe(-5)
    expect(parseTextToColumnsNumber('5-')).toBeNull()
    expect(
      parseTextToColumnsNumber('7', { decimal: '.', thousands: '.', trailingMinus: false }),
    ).toBeNull()
    expect(textToColumnsCell('1.234,5', 'general', european)).toMatchObject({ v: 1234.5 })
    expect(
      planTextToColumns(
        seed,
        { ...comma, numbers: { decimal: '.', thousands: '.', trailingMinus: false } },
        { row: 4, column: 0 },
        [],
        () => 'free',
      ).error,
    ).toBe('numbers')
  })

  it('classifies destination cells so Excel-style overwrite needs confirmation', () => {
    const plan = planTextToColumns(
      seed,
      comma,
      { row: 4, column: 0 },
      ['skip', 'general', 'general'],
      (row, column) => (column === 1 && row === 5 ? 'overwrite' : 'free'),
    )
    expect(plan.plan?.overwrites).toEqual([{ row: 5, column: 1 }])
    // a skipped column never counts as overwritten
    expect(
      planTextToColumns(
        seed,
        comma,
        { row: 4, column: 0 },
        ['general', 'skip', 'general'],
        () => 'overwrite',
      ).plan?.overwrites.map((cell) => cell.column),
    ).toEqual([2, 2])
  })

  it('reproduces the Excel Mac fixed-width specimen split at characters 4 and 7', async () => {
    const directory = join(
      fileURLToPath(new URL('../../..', import.meta.url)),
      'docs/qa/office-mac-field-evidence-2026-10-13/sheets-fixed-width',
    )
    const after = await JSZip.loadAsync(
      await readFile(join(directory, 'excel-mac-fixed-width-after.xlsx')),
    )
    const sheet = await after.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<f>10+5</f>')
    expect(splitFixedWidthCell('ABCD001XYZ', [4, 7])).toEqual({ fields: ['ABCD', '001', 'XYZ'] })
    expect(splitFixedWidthCell('ABCD00', [4, 7])).toEqual({ fields: ['ABCD', '00', ''] })
  })
})
