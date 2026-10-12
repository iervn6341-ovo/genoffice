/// Home → Fill → Series… (Excel's Series dialog). Pure planning: given the selection, the options
/// and a reader for the numeric value of a cell, it returns the cells to write. The behavior was
/// observed in Excel 16 through COM (`Range.DataSeries`):
///   - each column (Series in Columns) or row (Series in Rows) starts from its own first cell;
///   - Linear adds the step, Growth multiplies by it, Date moves by a day / weekday / month / year;
///   - the Stop value ends the series (inclusive) and leaves the cells after it untouched;
///   - a single selected cell with a Stop value extends past the selection up to the stop;
///   - Trend replaces the cells with the least-squares line (Linear) or exponential curve (Growth)
///     through the numbers already there, and ignores Step and Stop;
///   - a line whose first cell is not a number is left alone.

export type SeriesAxis = 'rows' | 'columns'
export type SeriesType = 'linear' | 'growth' | 'date'
export type SeriesDateUnit = 'day' | 'weekday' | 'month' | 'year'

export interface SeriesOptions {
  readonly axis: SeriesAxis
  readonly type: SeriesType
  readonly dateUnit: SeriesDateUnit
  readonly step: number
  /// null = no stop value
  readonly stop: number | null
  readonly trend: boolean
  /// workbook uses the 1904 date system
  readonly date1904?: boolean
}

export interface SeriesRect {
  readonly startRow: number
  readonly startColumn: number
  readonly rowCount: number
  readonly columnCount: number
}

export interface SeriesLimits {
  readonly maxRows: number
  readonly maxColumns: number
}

export interface SeriesWrite {
  readonly row: number
  readonly column: number
  readonly value: number
}

export type SeriesSkip =
  'no-numeric-start' | 'stop-required' | 'too-large' | 'trend-needs-positive' | 'invalid-number'

export interface SeriesPlan {
  readonly writes: readonly SeriesWrite[]
  /// the first cell of each line that produced writes: Excel copies its format onto the line
  readonly starts: readonly { readonly row: number; readonly column: number }[]
  /// why nothing was written (null when something was)
  readonly skipped: SeriesSkip | null
}

const DAY_MS = 86_400_000

function epochMs(date1904: boolean): number {
  return date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
}

export function serialToYmd(
  serial: number,
  date1904 = false,
): { year: number; month: number; day: number } {
  const date = new Date(epochMs(date1904) + Math.floor(serial) * DAY_MS)
  return { year: date.getUTCFullYear(), month: date.getUTCMonth(), day: date.getUTCDate() }
}

export function ymdToSerial(year: number, month: number, day: number, date1904 = false): number {
  return Math.round((Date.UTC(year, month, day) - epochMs(date1904)) / DAY_MS)
}

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month + 1, 0)).getUTCDate()
}

/// `start` moved by `months` calendar months; the day clamps to the month's last day (Jan 31 +1 = Feb 29)
export function addMonths(serial: number, months: number, date1904 = false): number {
  const { year, month, day } = serialToYmd(serial, date1904)
  const total = year * 12 + month + months
  const targetYear = Math.floor(total / 12)
  const targetMonth = total - targetYear * 12
  return ymdToSerial(
    targetYear,
    targetMonth,
    Math.min(day, daysInMonth(targetYear, targetMonth)),
    date1904,
  )
}

function isWeekend(serial: number, date1904: boolean): boolean {
  const day = new Date(epochMs(date1904) + Math.floor(serial) * DAY_MS).getUTCDay()
  return day === 0 || day === 6
}

/// `serial` moved by `count` working days (Saturday and Sunday skipped), like Excel's Weekday series
export function addWeekdays(serial: number, count: number, date1904 = false): number {
  let current = Math.floor(serial)
  const direction = count < 0 ? -1 : 1
  for (let moved = 0; moved < Math.abs(count); moved += 1) {
    do current += direction
    while (isWeekend(current, date1904))
  }
  return current
}

function dateValueAt(start: number, index: number, opts: SeriesOptions): number {
  const date1904 = opts.date1904 === true
  switch (opts.dateUnit) {
    case 'day':
      return start + index * opts.step
    case 'weekday':
      return addWeekdays(start, index * opts.step, date1904)
    case 'month':
      return addMonths(start, index * opts.step, date1904)
    case 'year':
      return addMonths(start, index * opts.step * 12, date1904)
  }
}

/// 15 significant digits, like the cell Excel stores (1.1 + 0.1 must read 1.2)
const tidy = (value: number): number => Number(value.toPrecision(15))

function valueAt(start: number, index: number, opts: SeriesOptions): number {
  if (opts.type === 'growth') return tidy(start * opts.step ** index)
  if (opts.type === 'date') return dateValueAt(start, index, opts)
  return tidy(start + index * opts.step)
}

/**
 * The series' value at an index, for indexes asked in order (the planner walks them 1, 2, 3 …). A
 * weekday series steps from the previous value: from the start every time it would be quadratic and
 * freeze the window on a long series.
 */
function makeValueAt(start: number, opts: SeriesOptions): (index: number) => number {
  if (opts.type !== 'date' || opts.dateUnit !== 'weekday') {
    return (index) => valueAt(start, index, opts)
  }
  const date1904 = opts.date1904 === true
  let lastIndex = 0
  let lastValue = Math.floor(start)
  return (index) => {
    if (index === lastIndex) return lastValue
    lastValue =
      index === lastIndex + 1
        ? addWeekdays(lastValue, opts.step, date1904)
        : addWeekdays(start, index * opts.step, date1904)
    lastIndex = index
    return lastValue
  }
}

/// Least squares through (index, value) points; null when the points do not pin a line down
function fitLine(points: readonly (readonly [number, number])[]): { a: number; b: number } | null {
  if (points.length < 2) return null
  const n = points.length
  let sx = 0
  let sy = 0
  let sxy = 0
  let sxx = 0
  for (const [x, y] of points) {
    sx += x
    sy += y
    sxy += x * y
    sxx += x * x
  }
  const denominator = n * sxx - sx * sx
  if (denominator === 0) return null
  const b = (n * sxy - sx * sy) / denominator
  return { a: (sy - b * sx) / n, b }
}

interface Line {
  readonly cells: readonly { readonly row: number; readonly column: number }[]
}

function linesOf(rect: SeriesRect, axis: SeriesAxis): Line[] {
  const lines: Line[] = []
  if (axis === 'columns') {
    for (let c = 0; c < rect.columnCount; c += 1) {
      lines.push({
        cells: Array.from({ length: rect.rowCount }, (_, r) => ({
          row: rect.startRow + r,
          column: rect.startColumn + c,
        })),
      })
    }
  } else {
    for (let r = 0; r < rect.rowCount; r += 1) {
      lines.push({
        cells: Array.from({ length: rect.columnCount }, (_, c) => ({
          row: rect.startRow + r,
          column: rect.startColumn + c,
        })),
      })
    }
  }
  return lines
}

/// Did `value` pass `stop` going the way the series is heading?
function pastStop(value: number, stop: number, heading: number): boolean {
  if (heading > 0) return value > stop
  if (heading < 0) return value < stop
  return false
}

/// How many cells a series needs to reach `stop`, counting the start; null when it never would
function lengthToStop(
  start: number,
  stop: number,
  opts: SeriesOptions,
  cap: number,
): number | null {
  const at = makeValueAt(start, opts)
  const heading = Math.sign(at(1) - start)
  if (heading === 0 || pastStop(start, stop, heading)) return heading === 0 ? null : 1
  let count = 1
  while (count < cap + 1) {
    if (pastStop(at(count), stop, heading)) return count
    count += 1
  }
  return cap + 1
}

export function planSeries(
  rect: SeriesRect,
  opts: SeriesOptions,
  readNumber: (row: number, column: number) => number | null,
  limits: SeriesLimits,
): SeriesPlan {
  if (!Number.isFinite(opts.step) || (opts.stop !== null && !Number.isFinite(opts.stop))) {
    return { writes: [], starts: [], skipped: 'invalid-number' }
  }
  const writes: SeriesWrite[] = []
  const starts: { row: number; column: number }[] = []
  let skipped: SeriesSkip | null = null
  const singleCell = rect.rowCount === 1 && rect.columnCount === 1
  const trend = opts.trend && opts.type !== 'date'

  for (const line of linesOf(rect, opts.axis)) {
    const first = line.cells[0]
    if (!first) continue
    const start = readNumber(first.row, first.column)
    if (start === null) {
      skipped ??= 'no-numeric-start'
      continue
    }

    if (trend) {
      const points: [number, number][] = []
      line.cells.forEach((cell, index) => {
        const value = readNumber(cell.row, cell.column)
        if (value !== null) points.push([index, value])
      })
      if (opts.type === 'growth' && points.some(([, y]) => y <= 0)) {
        skipped ??= 'trend-needs-positive'
        continue
      }
      const fit = fitLine(
        opts.type === 'growth' ? points.map(([x, y]) => [x, Math.log(y)] as const) : points,
      )
      line.cells.forEach((cell, index) => {
        const value = fit
          ? opts.type === 'growth'
            ? Math.exp(fit.a + fit.b * index)
            : fit.a + fit.b * index
          : valueAt(start, index, opts) // one number: nothing to fit, step on from it
        writes.push({ row: cell.row, column: cell.column, value: tidy(value) })
      })
      starts.push(first)
      continue
    }

    let length = line.cells.length
    if (singleCell) {
      // one cell: the Stop value says how far to go; without it there is nothing to do
      if (opts.stop === null) {
        skipped ??= 'stop-required'
        continue
      }
      const room =
        opts.axis === 'columns' ? limits.maxRows - first.row : limits.maxColumns - first.column
      const needed = lengthToStop(start, opts.stop, opts, room)
      if (needed === null) continue
      if (needed > room) {
        skipped ??= 'too-large'
        continue
      }
      length = needed
    }
    const at = makeValueAt(start, opts)
    const heading = Math.sign(at(1) - start)
    for (let index = 1; index < length; index += 1) {
      const value = at(index)
      if (opts.stop !== null && pastStop(value, opts.stop, heading)) break
      const cell =
        opts.axis === 'columns'
          ? { row: first.row + index, column: first.column }
          : { row: first.row, column: first.column + index }
      writes.push({ row: cell.row, column: cell.column, value })
    }
    starts.push(first)
  }
  return { writes, starts, skipped: writes.length > 0 || starts.length > 0 ? null : skipped }
}

/// the dialog's "Series in" preselection: a single row runs along the row, anything else down
export function defaultSeriesAxis(rect: SeriesRect): SeriesAxis {
  return rect.rowCount === 1 && rect.columnCount > 1 ? 'rows' : 'columns'
}
