import { getSystemShortDate } from '@genoffice/xlsx-gateway/shared/short-date'
import { CellValueType, type ICellData } from '@univerjs/core'

/** Excel's Advanced Text Import Settings (decimal / thousands separators, trailing minus). */
export interface TextToColumnsNumberSettings {
  readonly decimal: string
  readonly thousands: string
  readonly trailingMinus: boolean
}

export const DEFAULT_NUMBER_SETTINGS: TextToColumnsNumberSettings = {
  decimal: '.',
  thousands: ',',
  trailingMinus: false,
}

export interface TextToColumnsOptions {
  readonly fixedWidth?: { readonly breaks: readonly number[] }
  readonly numbers?: TextToColumnsNumberSettings
  readonly tab: boolean
  readonly semicolon: boolean
  readonly comma: boolean
  readonly space: boolean
  readonly other: string
  readonly collapse: boolean
  readonly qualifier: 'double' | 'none'
}

export type TextToColumnsDateOrder = 'mdy' | 'dmy' | 'ymd' | 'myd' | 'dym' | 'ydm'

export const TEXT_TO_COLUMNS_DATE_ORDERS: readonly TextToColumnsDateOrder[] = [
  'mdy',
  'dmy',
  'ymd',
  'myd',
  'dym',
  'ydm',
]

export type TextToColumnsFormat = 'general' | 'text' | 'skip' | `date-${TextToColumnsDateOrder}`

export function textToColumnsDateOrder(format: TextToColumnsFormat): TextToColumnsDateOrder | null {
  return format.startsWith('date-') ? (format.slice(5) as TextToColumnsDateOrder) : null
}

export interface TextToColumnsSourceRow {
  readonly row: number
  readonly text: string
}

export interface TextToColumnsSeed {
  readonly workbookId: string
  readonly sheetId: string
  readonly sourceColumn: number
  readonly rows: readonly TextToColumnsSourceRow[]
  readonly maxRows: number
  readonly maxColumns: number
  /** 1904 date system workbooks need a different serial epoch for Date columns. */
  readonly date1904?: boolean
}

export type TextToColumnsError =
  'delimiter' | 'breaks' | 'qualifier' | 'source' | 'destination' | 'occupied' | 'size' | 'numbers'

/**
 * What the destination cell holds: `free` is empty, `overwrite` is plain data or a simple
 * formula that Excel replaces after a prompt, `blocked` is a merged cell, array/shared
 * formula or rich text that this wizard never rewrites.
 */
export type TextToColumnsOccupancy = 'free' | 'overwrite' | 'blocked'

export interface TextToColumnsColumnWrite {
  readonly column: number
  readonly values: readonly (readonly ICellData[])[]
}

export interface TextToColumnsPlan {
  readonly startRow: number
  readonly writes: readonly TextToColumnsColumnWrite[]
  readonly preview: readonly (readonly string[])[]
  /** Destination cells holding data that the write replaces (needs the user's confirmation). */
  readonly overwrites: readonly { readonly row: number; readonly column: number }[]
}

export type TextToColumnsResult =
  | { readonly plan: TextToColumnsPlan; readonly error?: never }
  | { readonly error: TextToColumnsError; readonly plan?: never }

/** Excel's text qualifier is recognized only at the beginning of a field. */
export function splitDelimitedCell(
  text: string,
  options: TextToColumnsOptions,
): { fields: string[]; error?: never } | { error: TextToColumnsError; fields?: never } {
  const delimiters = [
    options.tab ? '\t' : '',
    options.semicolon ? ';' : '',
    options.comma ? ',' : '',
    options.space ? ' ' : '',
    options.other,
  ].filter(Boolean)
  if (options.other && [...options.other].length !== 1) return { error: 'delimiter' }
  if (delimiters.length === 0) return { error: 'delimiter' }
  const separators = new Set(delimiters)
  const fields: string[] = []
  let field = ''
  let quoted = false
  let closed = false
  let atStart = true
  let previousWasSeparator = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]!
    if (options.qualifier === 'double' && character === '"') {
      if (quoted && text[index + 1] === '"') {
        field += '"'
        index += 1
        continue
      }
      if (quoted) {
        quoted = false
        closed = true
        continue
      }
      if (atStart) {
        quoted = true
        atStart = false
        continue
      }
      if (closed) return { error: 'qualifier' }
    }
    if (!quoted && separators.has(character)) {
      if (!options.collapse || !previousWasSeparator) fields.push(field)
      field = ''
      atStart = true
      closed = false
      previousWasSeparator = true
      continue
    }
    if (closed) return { error: 'qualifier' }
    field += character
    atStart = false
    previousWasSeparator = false
  }
  if (quoted) return { error: 'qualifier' }
  if (!options.collapse || !previousWasSeparator || fields.length === 0) fields.push(field)
  return { fields }
}

/** Fixed-width breaks are one-based Unicode character positions between columns. */
export function splitFixedWidthCell(
  text: string,
  breaks: readonly number[],
): { fields: string[]; error?: never } | { error: TextToColumnsError; fields?: never } {
  if (
    breaks.length === 0 ||
    breaks.length > 15 ||
    breaks.some(
      (position, index) =>
        !Number.isInteger(position) ||
        position < 1 ||
        position > 32_767 ||
        (index > 0 && position <= breaks[index - 1]!),
    )
  )
    return { error: 'breaks' }
  const characters = [...text]
  const fields: string[] = []
  let start = 0
  for (const position of breaks) {
    fields.push(characters.slice(start, position).join(''))
    start = position
  }
  fields.push(characters.slice(start).join(''))
  return { fields }
}

const NUMERIC = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function validNumberSettings(settings: TextToColumnsNumberSettings): boolean {
  const decimal = [...settings.decimal]
  const thousands = [...settings.thousands]
  return (
    decimal.length === 1 &&
    thousands.length <= 1 &&
    settings.decimal !== settings.thousands &&
    !/[\d+\-eE\s]/.test(settings.decimal) &&
    !/[\d+\-eE\s]/.test(settings.thousands)
  )
}

/** General-format number recognition honouring Advanced separators; null when text. */
export function parseTextToColumnsNumber(
  value: string,
  settings: TextToColumnsNumberSettings = DEFAULT_NUMBER_SETTINGS,
): number | null {
  let text = value.trim()
  if (text === '' || !validNumberSettings(settings)) return null
  if (settings.trailingMinus && /^[^+-].*-$/.test(text)) text = `-${text.slice(0, -1)}`
  const sign = /^[+-]/.exec(text)?.[0] ?? ''
  let body = sign ? text.slice(1) : text
  const decimalAt = body.indexOf(settings.decimal)
  let integerPart = decimalAt >= 0 ? body.slice(0, decimalAt) : body
  const rest = decimalAt >= 0 ? body.slice(decimalAt) : ''
  if (settings.thousands && integerPart.includes(settings.thousands)) {
    const grouped = new RegExp(`^\\d{1,3}(?:${escapeRegExp(settings.thousands)}\\d{3})+$`)
    if (!grouped.test(integerPart)) return null
    integerPart = integerPart.split(settings.thousands).join('')
  }
  body = integerPart + (rest ? `.${rest.slice(settings.decimal.length)}` : '')
  const normalized = `${sign}${body}`
  if (!NUMERIC.test(normalized)) return null
  const number = Number(normalized)
  return Number.isFinite(number) ? number : null
}

const MONTH_NAMES = [
  'january',
  'february',
  'march',
  'april',
  'may',
  'june',
  'july',
  'august',
  'september',
  'october',
  'november',
  'december',
]

function monthFromName(token: string): number | null {
  const lower = token.toLowerCase().replace(/\.$/, '')
  if (lower.length < 3) return null
  const index = MONTH_NAMES.findIndex((name) => name === lower || name.slice(0, 3) === lower)
  return index >= 0 ? index + 1 : null
}

/** Excel's two-digit year window: 00–29 → 2000s, 30–99 → 1900s. */
function expandYear(token: string): number {
  const year = Number(token)
  if (token.length > 2) return year
  return year < 30 ? 2000 + year : 1900 + year
}

/**
 * Parse a date typed in the given Y/M/D order into a workbook serial, or null when Excel
 * would leave the text alone. Accepts `/`, `-`, `.` or space separators, English month names
 * and the compact 6/8-digit forms (e.g. 20241213 for YMD).
 */
export function parseTextToColumnsDate(
  value: string,
  order: TextToColumnsDateOrder,
  date1904 = false,
): number | null {
  const text = value.trim()
  let tokens: string[]
  const compact = /^\d{6}$|^\d{8}$/.test(text)
  if (compact) {
    const yearLength = text.length === 8 ? 4 : 2
    const lengths = [...order].map((part) => (part === 'y' ? yearLength : 2))
    tokens = []
    let at = 0
    for (const length of lengths) {
      tokens.push(text.slice(at, at + length))
      at += length
    }
  } else {
    tokens = text.split(/[\s/\-.,]+/).filter(Boolean)
    if (tokens.length !== 3 || !/^[\p{L}\d.]+$/u.test(tokens.join(''))) return null
  }
  if (tokens.length !== 3) return null
  let year: number | null = null
  let month: number | null = null
  let day: number | null = null
  const named = tokens.findIndex((token) => /^\p{L}/u.test(token))
  if (named < 0) {
    for (let index = 0; index < 3; index += 1) {
      const token = tokens[index]!
      const part = order[index]!
      if (!/^\d{1,4}$/.test(token)) return null
      if (part === 'y') year = expandYear(token)
      else if (part === 'm') month = Number(token)
      else day = Number(token)
    }
  } else {
    // a month name fixes the month; the remaining two numbers keep their relative order
    month = monthFromName(tokens[named]!)
    const numbers = tokens.filter((_, index) => index !== named)
    const remaining = [...order].filter((part) => part !== 'm')
    if (month === null || numbers.some((token) => !/^\d{1,4}$/.test(token))) return null
    for (let index = 0; index < 2; index += 1) {
      if (remaining[index] === 'y') year = expandYear(numbers[index]!)
      else day = Number(numbers[index])
    }
  }
  if (year === null || month === null || day === null) return null
  if (month < 1 || month > 12 || day < 1 || year < 1900 || year > 9999) return null
  if (day > new Date(Date.UTC(year, month, 0)).getUTCDate()) return null
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  let serial = Math.round((Date.UTC(year, month - 1, day) - epoch) / 86_400_000)
  if (date1904 && serial < 0) return null
  // Excel's 1900 leap-year bug: serials before 1900-03-01 are one day higher than the epoch math
  if (!date1904 && serial < 61) serial -= 1
  return serial
}

export function textToColumnsCell(
  value: string,
  format: TextToColumnsFormat,
  numbers: TextToColumnsNumberSettings = DEFAULT_NUMBER_SETTINGS,
  date1904 = false,
): ICellData {
  if (format === 'general') {
    const number = parseTextToColumnsNumber(value, numbers)
    if (number !== null) return { v: number, t: CellValueType.NUMBER, f: null, p: null, si: null }
  }
  const order = textToColumnsDateOrder(format)
  if (order) {
    const serial = parseTextToColumnsDate(value, order, date1904)
    if (serial !== null)
      return {
        v: serial,
        t: CellValueType.NUMBER,
        f: null,
        p: null,
        si: null,
        s: { n: { pattern: getSystemShortDate() } },
      }
  }
  return {
    v: value,
    t: format === 'text' ? CellValueType.FORCE_STRING : CellValueType.STRING,
    f: null,
    p: null,
    si: null,
  }
}

/** Plan every write before touching Univer; callers recheck the source before applying. */
export function planTextToColumns(
  seed: TextToColumnsSeed,
  options: TextToColumnsOptions,
  destination: { row: number; column: number },
  formats: readonly TextToColumnsFormat[],
  occupied: (row: number, column: number) => TextToColumnsOccupancy,
): TextToColumnsResult {
  if (seed.rows.length === 0 || seed.rows.length > 1_000) return { error: 'size' }
  if (options.numbers && !validNumberSettings(options.numbers)) return { error: 'numbers' }
  if (
    !Number.isInteger(destination.row) ||
    !Number.isInteger(destination.column) ||
    destination.row < 0 ||
    destination.column < 0 ||
    destination.row + seed.rows.length > seed.maxRows
  )
    return { error: 'destination' }
  const preview: string[][] = []
  let width = 0
  for (const source of seed.rows) {
    if (source.text.length > 32_767) return { error: 'size' }
    const parsed = options.fixedWidth
      ? splitFixedWidthCell(source.text, options.fixedWidth.breaks)
      : splitDelimitedCell(source.text, options)
    if (parsed.error) return { error: parsed.error }
    preview.push(parsed.fields)
    width = Math.max(width, parsed.fields.length)
  }
  if (width === 0 || width > 16 || width * seed.rows.length > 10_000) return { error: 'size' }
  if (destination.column + width > seed.maxColumns) return { error: 'destination' }
  const sourceFirst = seed.rows[0]!.row
  const sourceLast = seed.rows[seed.rows.length - 1]!.row
  const overlapsSource =
    destination.column <= seed.sourceColumn &&
    destination.column + width > seed.sourceColumn &&
    destination.row <= sourceLast &&
    destination.row + seed.rows.length - 1 >= sourceFirst
  if (
    overlapsSource &&
    (destination.column !== seed.sourceColumn || destination.row !== sourceFirst)
  )
    return { error: 'destination' }
  const writes: TextToColumnsColumnWrite[] = []
  const overwrites: { row: number; column: number }[] = []
  for (let offset = 0; offset < width; offset += 1) {
    if ((formats[offset] ?? 'general') === 'skip') continue
    const column = destination.column + offset
    const isSourceColumn = column === seed.sourceColumn && destination.row === sourceFirst
    for (let rowOffset = 0; rowOffset < seed.rows.length; rowOffset += 1) {
      if (isSourceColumn) continue
      const state = occupied(destination.row + rowOffset, column)
      if (state === 'blocked') return { error: 'occupied' }
      if (state === 'overwrite') overwrites.push({ row: destination.row + rowOffset, column })
    }
    writes.push({
      column,
      values: preview.map((row) => [
        textToColumnsCell(
          row[offset] ?? '',
          formats[offset] ?? 'general',
          options.numbers,
          seed.date1904,
        ),
      ]),
    })
  }
  if (writes.length === 0) return { error: 'source' }
  return { plan: { startRow: destination.row, writes, preview, overwrites } }
}

export function parseTextToColumnsAddress(value: string): { row: number; column: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(value.trim())
  if (!match) return null
  let column = 0
  for (const character of match[1]!.toUpperCase())
    column = column * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

export function textToColumnsAddress(row: number, column: number): string {
  let current = column + 1
  let label = ''
  while (current > 0) {
    current -= 1
    label = String.fromCharCode(65 + (current % 26)) + label
    current = Math.floor(current / 26)
  }
  return `${label}${row + 1}`
}
