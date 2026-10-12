import { describe, expect, it } from 'vitest'
import {
  addMonths,
  addWeekdays,
  defaultSeriesAxis,
  planSeries,
  serialToYmd,
  ymdToSerial,
  type SeriesOptions,
  type SeriesRect,
} from '../src/renderer/fill-series'

// Every expectation below was read from Excel 16 (Range.DataSeries via COM), see
// docs/qa/handoff-ribbon-audit-2026-10-11.md.

const LIMITS = { maxRows: 1000, maxColumns: 100 }

const base: SeriesOptions = {
  axis: 'columns',
  type: 'linear',
  dateUnit: 'day',
  step: 1,
  stop: null,
  trend: false,
}

/** a sheet of numbers keyed "row:col" */
function sheet(cells: Record<string, number>) {
  return (row: number, column: number): number | null => cells[`${row}:${column}`] ?? null
}

/** applies a plan onto a copy of the sheet and reads one column back */
function column(
  rect: SeriesRect,
  opts: SeriesOptions,
  cells: Record<string, number>,
  rows: number,
  col = 0,
) {
  const plan = planSeries(rect, opts, sheet(cells), LIMITS)
  const grid = { ...cells }
  for (const write of plan.writes) grid[`${write.row}:${write.column}`] = write.value
  return Array.from({ length: rows }, (_, r) => grid[`${r}:${col}`] ?? null)
}

const A1_A5: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 5, columnCount: 1 }

describe('linear and growth series (Excel DataSeries)', () => {
  it('linear start 1 step 2 down A1:A5 → 1 3 5 7 9', () => {
    expect(column(A1_A5, { ...base, step: 2 }, { '0:0': 1 }, 5)).toEqual([1, 3, 5, 7, 9])
  })

  it('growth start 2 step 3 → 2 6 18 54 162', () => {
    expect(column(A1_A5, { ...base, type: 'growth', step: 3 }, { '0:0': 2 }, 5)).toEqual([
      2, 6, 18, 54, 162,
    ])
  })

  it('negative step along a row: 5 4 3 2 1', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 1, columnCount: 5 }
    const plan = planSeries(rect, { ...base, axis: 'rows', step: -1 }, sheet({ '0:0': 5 }), LIMITS)
    expect(plan.writes.map((w) => [w.column, w.value])).toEqual([
      [1, 4],
      [2, 3],
      [3, 2],
      [4, 1],
    ])
  })

  it('stop value ends the series (inclusive) and leaves later cells alone', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 10, columnCount: 1 }
    expect(column(rect, { ...base, step: 2, stop: 7 }, { '0:0': 1 }, 6)).toEqual([
      1,
      3,
      5,
      7,
      null,
      null,
    ])
    // the 99 after the stop survives
    expect(
      column({ ...A1_A5, rowCount: 6 }, { ...base, stop: 3 }, { '0:0': 1, '4:0': 99 }, 6),
    ).toEqual([1, 2, 3, null, 99, null])
  })

  it('growth stops at the stop value: 2 4 8 16 for stop 20', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 10, columnCount: 1 }
    expect(column(rect, { ...base, type: 'growth', step: 2, stop: 20 }, { '0:0': 2 }, 6)).toEqual([
      2,
      4,
      8,
      16,
      null,
      null,
    ])
  })

  it('a shrinking growth series stops below the stop: 4 2 1 0.5 0.25 for stop 0.2', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 10, columnCount: 1 }
    expect(
      column(rect, { ...base, type: 'growth', step: 0.5, stop: 0.2 }, { '0:0': 4 }, 6),
    ).toEqual([4, 2, 1, 0.5, 0.25, null])
  })

  it('descending linear stops below the stop: 10 7 4 for step −3 stop 2', () => {
    expect(
      column({ ...A1_A5, rowCount: 6 }, { ...base, step: -3, stop: 2 }, { '0:0': 10 }, 6),
    ).toEqual([10, 7, 4, null, null, null])
  })

  it('a stop below the start with a positive step writes nothing', () => {
    expect(column(A1_A5, { ...base, stop: 0 }, { '0:0': 5 }, 5)).toEqual([
      5,
      null,
      null,
      null,
      null,
    ])
  })

  it('step 0 repeats the start', () => {
    expect(column({ ...A1_A5, rowCount: 4 }, { ...base, step: 0 }, { '0:0': 3 }, 4)).toEqual([
      3, 3, 3, 3,
    ])
  })

  it('negative growth: −2 −4 −8 −16', () => {
    expect(
      column({ ...A1_A5, rowCount: 4 }, { ...base, type: 'growth', step: 2 }, { '0:0': -2 }, 4),
    ).toEqual([-2, -4, -8, -16])
  })

  it('decimal steps do not drift: 1 1.1 1.2 1.3', () => {
    expect(column({ ...A1_A5, rowCount: 4 }, { ...base, step: 0.1 }, { '0:0': 1 }, 4)).toEqual([
      1, 1.1, 1.2, 1.3,
    ])
  })

  it('every column starts from its own first cell', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 3, columnCount: 2 }
    const cells = { '0:0': 1, '0:1': 10 }
    expect(column(rect, base, cells, 3, 0)).toEqual([1, 2, 3])
    expect(column(rect, base, cells, 3, 1)).toEqual([10, 11, 12])
  })

  it('every row starts from its own first cell when the series runs in rows', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 2, columnCount: 3 }
    const plan = planSeries(
      rect,
      { ...base, axis: 'rows', step: 5 },
      sheet({ '0:0': 1, '1:0': 100 }),
      LIMITS,
    )
    const read = (r: number, c: number) =>
      plan.writes.find((w) => w.row === r && w.column === c)?.value
    expect([read(0, 1), read(0, 2), read(1, 1), read(1, 2)]).toEqual([6, 11, 105, 110])
  })

  it('overwrites the other cells of the line', () => {
    expect(column({ ...A1_A5, rowCount: 3 }, base, { '0:0': 1, '1:0': 99, '2:0': 98 }, 3)).toEqual([
      1, 2, 3,
    ])
  })

  it('leaves a line alone when its first cell is not a number', () => {
    const plan = planSeries(A1_A5, base, sheet({ '1:0': 4 }), LIMITS)
    expect(plan.writes).toEqual([])
    expect(plan.skipped).toBe('no-numeric-start')
  })
})

describe('single cell + stop value', () => {
  const one: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 1, columnCount: 1 }

  it('extends down to the stop: A1=1 stop 5 → A1:A5', () => {
    expect(column(one, { ...base, stop: 5 }, { '0:0': 1 }, 8)).toEqual([
      1,
      2,
      3,
      4,
      5,
      null,
      null,
      null,
    ])
  })

  it('extends along a row for Series in Rows', () => {
    const plan = planSeries(one, { ...base, axis: 'rows', stop: 4 }, sheet({ '0:0': 1 }), LIMITS)
    expect(plan.writes.map((w) => [w.row, w.column, w.value])).toEqual([
      [0, 1, 2],
      [0, 2, 3],
      [0, 3, 4],
    ])
  })

  it('needs a stop value', () => {
    const plan = planSeries(one, base, sheet({ '0:0': 1 }), LIMITS)
    expect(plan.writes).toEqual([])
    expect(plan.skipped).toBe('stop-required')
  })

  it('refuses a series that would run off the sheet', () => {
    const plan = planSeries(one, { ...base, stop: 5000 }, sheet({ '0:0': 1 }), LIMITS)
    expect(plan.writes).toEqual([])
    expect(plan.skipped).toBe('too-large')
  })
})

describe('trend (Excel ignores Step and Stop)', () => {
  it('exact line 1,3 → 1 3 5 7 9', () => {
    expect(column(A1_A5, { ...base, trend: true }, { '0:0': 1, '1:0': 3 }, 5)).toEqual([
      1, 3, 5, 7, 9,
    ])
  })

  it('least squares replaces the cells: 1,2,4 → 0.8333 2.3333 3.8333 …', () => {
    const out = column(A1_A5, { ...base, trend: true }, { '0:0': 1, '1:0': 2, '2:0': 4 }, 5)
    const expected = [
      0.833333333333334, 2.33333333333333, 3.83333333333333, 5.33333333333333, 6.83333333333333,
    ]
    out.forEach((value, i) => expect(value).toBeCloseTo(expected[i]!, 10))
  })

  it('uses the positions of the numbers: 1 _ 4 → 1 2.5 4 5.5 7', () => {
    expect(column(A1_A5, { ...base, trend: true }, { '0:0': 1, '2:0': 4 }, 5)).toEqual([
      1, 2.5, 4, 5.5, 7,
    ])
  })

  it('growth trend 2,6 → 2 6 18 54 162; 2,4,9 → 1.96 4.16 8.83 …', () => {
    expect(
      column(A1_A5, { ...base, type: 'growth', trend: true }, { '0:0': 2, '1:0': 6 }, 5).map((v) =>
        Math.round(v as number),
      ),
    ).toEqual([2, 6, 18, 54, 162])
    const out = column(
      A1_A5,
      { ...base, type: 'growth', trend: true },
      { '0:0': 2, '1:0': 4, '2:0': 9 },
      5,
    )
    const expected = [
      1.96112183562192, 4.16016764610381, 8.82504826029864, 18.7207544074671, 39.7127171713439,
    ]
    out.forEach((value, i) => expect(value).toBeCloseTo(expected[i]!, 8))
  })

  it('one number alone steps on from it: 5 → 5 6 7 8', () => {
    expect(column({ ...A1_A5, rowCount: 4 }, { ...base, trend: true }, { '0:0': 5 }, 4)).toEqual([
      5, 6, 7, 8,
    ])
  })

  it('growth trend refuses non-positive numbers', () => {
    const plan = planSeries(
      A1_A5,
      { ...base, type: 'growth', trend: true },
      sheet({ '0:0': 2, '1:0': -1 }),
      LIMITS,
    )
    expect(plan.skipped).toBe('trend-needs-positive')
  })
})

describe('date series', () => {
  const dates = (n: number): SeriesRect => ({
    startRow: 0,
    startColumn: 0,
    rowCount: n,
    columnCount: 1,
  })
  const ymd = (serials: (number | null)[]) =>
    serials.map((s) => {
      if (s === null) return null
      const d = serialToYmd(s)
      return `${d.year}-${String(d.month + 1).padStart(2, '0')}-${String(d.day).padStart(2, '0')}`
    })

  it('days: 2024-01-30 + 1 → 01-31 02-01 02-02 02-03', () => {
    expect(
      ymd(column(dates(5), { ...base, type: 'date', dateUnit: 'day' }, { '0:0': 45321 }, 5)),
    ).toEqual(['2024-01-30', '2024-01-31', '2024-02-01', '2024-02-02', '2024-02-03'])
  })

  it('days with step 7', () => {
    expect(
      ymd(
        column(dates(4), { ...base, type: 'date', dateUnit: 'day', step: 7 }, { '0:0': 45321 }, 4),
      ),
    ).toEqual(['2024-01-30', '2024-02-06', '2024-02-13', '2024-02-20'])
  })

  it('weekdays skip Saturday and Sunday', () => {
    // Fri 2024-02-02
    expect(
      ymd(column(dates(6), { ...base, type: 'date', dateUnit: 'weekday' }, { '0:0': 45324 }, 6)),
    ).toEqual(['2024-02-02', '2024-02-05', '2024-02-06', '2024-02-07', '2024-02-08', '2024-02-09'])
  })

  it('weekday step 2 from Monday: Mon Wed Fri Tue', () => {
    expect(
      ymd(
        column(
          dates(4),
          { ...base, type: 'date', dateUnit: 'weekday', step: 2 },
          { '0:0': 45327 },
          4,
        ),
      ),
    ).toEqual(['2024-02-05', '2024-02-07', '2024-02-09', '2024-02-13'])
  })

  it('weekdays from a Saturday step to Monday', () => {
    expect(
      ymd(column(dates(4), { ...base, type: 'date', dateUnit: 'weekday' }, { '0:0': 45325 }, 4)),
    ).toEqual(['2024-02-03', '2024-02-05', '2024-02-06', '2024-02-07'])
  })

  it('months clamp to the month end from the START each time: Jan 31 → Feb 29 → Mar 31 → Apr 30', () => {
    expect(
      ymd(column(dates(5), { ...base, type: 'date', dateUnit: 'month' }, { '0:0': 45322 }, 5)),
    ).toEqual(['2024-01-31', '2024-02-29', '2024-03-31', '2024-04-30', '2024-05-31'])
  })

  it('months backwards', () => {
    expect(
      ymd(
        column(
          dates(4),
          { ...base, type: 'date', dateUnit: 'month', step: -1 },
          { '0:0': 45351 },
          4,
        ),
      ),
    ).toEqual(['2024-02-29', '2024-01-29', '2023-12-29', '2023-11-29'])
  })

  it('years: Feb 29 2024 + 2y → Feb 28 2026, Feb 29 2028', () => {
    expect(
      ymd(
        column(dates(4), { ...base, type: 'date', dateUnit: 'year', step: 2 }, { '0:0': 45351 }, 4),
      ),
    ).toEqual(['2024-02-29', '2026-02-28', '2028-02-29', '2030-02-28'])
  })

  it('a date stop is inclusive', () => {
    expect(
      ymd(
        column(
          dates(8),
          { ...base, type: 'date', dateUnit: 'day', stop: 45323 },
          { '0:0': 45321 },
          8,
        ),
      ),
    ).toEqual(['2024-01-30', '2024-01-31', '2024-02-01', null, null, null, null, null])
  })

  it('Date works on plain numbers too (Excel steps General 5 → 6 7 8)', () => {
    expect(column(dates(4), { ...base, type: 'date', dateUnit: 'day' }, { '0:0': 5 }, 4)).toEqual([
      5, 6, 7, 8,
    ])
  })

  it('the 1904 date system shifts the calendar', () => {
    expect(ymdToSerial(2024, 0, 30, true)).toBe(ymdToSerial(2024, 0, 30, false) - 1462)
    expect(serialToYmd(ymdToSerial(2024, 0, 30, true), true)).toEqual({
      year: 2024,
      month: 0,
      day: 30,
    })
  })

  it('addMonths / addWeekdays helpers', () => {
    expect(addMonths(45322, 1)).toBe(45351)
    expect(addWeekdays(45324, 1)).toBe(45327)
    expect(addWeekdays(45327, -1)).toBe(45324)
  })
})

describe('series axis default', () => {
  it('a single row runs along the row, everything else down', () => {
    expect(defaultSeriesAxis({ startRow: 0, startColumn: 0, rowCount: 1, columnCount: 4 })).toBe(
      'rows',
    )
    expect(defaultSeriesAxis({ startRow: 0, startColumn: 0, rowCount: 4, columnCount: 1 })).toBe(
      'columns',
    )
    expect(defaultSeriesAxis({ startRow: 0, startColumn: 0, rowCount: 3, columnCount: 3 })).toBe(
      'columns',
    )
    expect(defaultSeriesAxis({ startRow: 0, startColumn: 0, rowCount: 1, columnCount: 1 })).toBe(
      'columns',
    )
  })
})

describe('long series stay linear in the length', () => {
  it('a 100 000-row weekday series is planned in well under a second and stays correct', () => {
    const rect: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 100_000, columnCount: 1 }
    const began = Date.now()
    const plan = planSeries(
      rect,
      { ...base, type: 'date', dateUnit: 'weekday' },
      sheet({ '0:0': 45324 }),
      { maxRows: 1_000_000, maxColumns: 100 },
    )
    expect(Date.now() - began).toBeLessThan(2000)
    expect(plan.writes).toHaveLength(99_999)
    // Mon Feb 5 follows Fri Feb 2, then Tue, Wed …
    expect(plan.writes.slice(0, 3).map((w) => w.value)).toEqual([45327, 45328, 45329])
    expect(plan.writes[plan.writes.length - 1]!.value).toBe(addWeekdays(45324, 99_999))
  })

  it('a single cell with a far Stop value finds its length without recomputing from the start', () => {
    const one: SeriesRect = { startRow: 0, startColumn: 0, rowCount: 1, columnCount: 1 }
    const began = Date.now()
    const plan = planSeries(
      one,
      { ...base, type: 'date', dateUnit: 'weekday', stop: 45324 + 60_000 },
      sheet({ '0:0': 45324 }),
      { maxRows: 1_000_000, maxColumns: 100 },
    )
    expect(Date.now() - began).toBeLessThan(2000)
    expect(plan.writes.length).toBeGreaterThan(40_000)
  })
})
