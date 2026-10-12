import type { WorkbookTableAdd } from '../shared/desktop-api'

/**
 * Table Design tab: pure rules for the tables this session created (the ones the journal knows).
 * A file's own tables are view-only here, as with Format as Table.
 */
export interface TableDesignOptions {
  readonly bandedRows: boolean
  readonly bandedColumns: boolean
  readonly firstColumn: boolean
  readonly lastColumn: boolean
  /** Header Row: a table without one has no header band and no filter buttons */
  readonly headerRow: boolean
  /** Total Row: a row under the data with a label and a SUBTOTAL */
  readonly totalRow: boolean
  /** Filter Button: the buttons on the header row */
  readonly filterButton: boolean
}

/** what Format as Table gives a new table: banded rows, a header row and filter buttons (Excel's default) */
export const DEFAULT_TABLE_OPTIONS: TableDesignOptions = {
  bandedRows: true,
  bandedColumns: false,
  firstColumn: false,
  lastColumn: false,
  headerRow: true,
  totalRow: false,
  filterButton: true,
}

export function tableOptionsOf(entry: WorkbookTableAdd): TableDesignOptions {
  return {
    bandedRows: entry.bandedRows,
    bandedColumns: entry.bandedColumns === true,
    firstColumn: entry.firstColumn === true,
    lastColumn: entry.lastColumn === true,
    headerRow: entry.headerRow !== false,
    totalRow: entry.totalsRow === true,
    filterButton: entry.filterButton !== false,
  }
}

/** the journal fields for options: a flag at Excel's default is left out, as every older entry has it */
export function optionFields(
  options: Partial<TableDesignOptions>,
): Pick<
  WorkbookTableAdd,
  'bandedColumns' | 'firstColumn' | 'lastColumn' | 'headerRow' | 'totalsRow' | 'filterButton'
> {
  return {
    ...(options.bandedColumns ? { bandedColumns: true } : {}),
    ...(options.firstColumn ? { firstColumn: true } : {}),
    ...(options.lastColumn ? { lastColumn: true } : {}),
    ...(options.headerRow === false ? { headerRow: false } : {}),
    ...(options.totalRow ? { totalsRow: true } : {}),
    ...(options.filterButton === false ? { filterButton: false } : {}),
  }
}

export interface TableArea {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

/**
 * The range Univer's table covers for a journal `area` (which is the table as the file holds it).
 * Univer reads the first row of its table as the header and counts structured references from the
 * row below it, and it has no totals row: without a header row the table reaches up over the empty
 * row the header left (its names come from the journal), and the totals row stays outside so that
 * `Table1[Col]` does not include the cell that sums it.
 */
export function univerTableRange(
  area: TableArea,
  options: Pick<TableDesignOptions, 'headerRow' | 'totalRow'>,
): TableArea {
  return {
    ...area,
    startRow: options.headerRow ? area.startRow : area.startRow - 1,
    endRow: options.totalRow ? area.endRow - 1 : area.endRow,
  }
}

/** the first row of data and the last (a header row above, a totals row below, are not data) */
export function dataRowsOf(
  area: TableArea,
  options: Pick<TableDesignOptions, 'headerRow' | 'totalRow'>,
): { first: number; last: number } {
  return {
    first: area.startRow + (options.headerRow ? 1 : 0),
    last: area.endRow - (options.totalRow ? 1 : 0),
  }
}

/** the totals row's function for a column: Excel sums a column of numbers and counts anything else */
export type TotalsKind = 'sum' | 'count'

export function totalsKindOf(values: readonly unknown[]): TotalsKind {
  const filled = values.filter((value) => value !== null && value !== undefined && value !== '')
  return filled.length > 0 && filled.every((value) => typeof value === 'number') ? 'sum' : 'count'
}

/** SUBTOTAL's function number: 109 sums, 103 counts (COUNTA), as Excel writes them */
export const SUBTOTAL_CODE: Record<TotalsKind, number> = { sum: 109, count: 103 }

/** the formula Excel writes into a totals cell, with the structured reference of its column */
export function totalsFormula(kind: TotalsKind, tableName: string, columnName: string): string {
  // a special character in the name is escaped with a quote, as in Excel's own references
  const escaped = columnName.replace(/(['[\]#])/g, "'$1")
  return `=SUBTOTAL(${SUBTOTAL_CODE[kind]},${tableName}[${escaped}])`
}

/** the session table that contains the cell, if any */
export function sessionTableAt(
  tables: readonly WorkbookTableAdd[],
  sheetId: string,
  row: number,
  column: number,
): WorkbookTableAdd | null {
  return (
    tables.find(
      (table) =>
        table.sheetId === sheetId &&
        row >= table.area.startRow &&
        row <= table.area.endRow &&
        column >= table.area.startColumn &&
        column <= table.area.endColumn,
    ) ?? null
  )
}

export type TableNameProblem = 'empty' | 'invalid' | 'taken' | 'reserved'

/**
 * Excel's table name rules, held to the stricter ones Univer's table service enforces (it refuses
 * a name outright, or renames it, and the saved file would then disagree with the grid): starts with
 * a letter or underscore (not a backslash); then letters, digits, underscores or periods; no spaces
 * or \ / : ? * [ ]; not a cell reference ($A1, AB12), not R1C1, R or C, not purely numeric; at most
 * 255 characters; not another table's name (any case) and not a name `reserved` by a sheet or a
 * defined name.
 */
export function tableNameProblem(
  name: string,
  taken: readonly string[],
  own?: string,
  reserved: readonly string[] = [],
): TableNameProblem | null {
  if (name.length === 0) return 'empty'
  if (name.length > 255) return 'invalid'
  if (!/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name)) return 'invalid'
  if (/^\$?[A-Za-z]{1,3}\$?\d+$/.test(name)) return 'invalid'
  if (/^[Rr]\d*[Cc]\d*$/.test(name) || /^[RrCc]$/.test(name)) return 'invalid'
  const lower = name.toLowerCase()
  if (taken.some((other) => other.toLowerCase() === lower && other !== own)) return 'taken'
  if (reserved.some((other) => other.toLowerCase() === lower)) return 'reserved'
  return null
}

/** the widest table Resize Table accepts (Excel's own limit is the sheet's 16 384 columns) */
const MAX_TABLE_COLUMNS = 1000

// ── Header Row / Total Row / Resize: what changes in the cells, decided without the grid ────────

export interface CellWrite {
  readonly row: number
  readonly column: number
  /** a text, a formula (starts with =), or null to empty the cell */
  readonly value: string | null
}

export type TableChangeError =
  'noRoomAbove' | 'noRoomBelow' | 'needsData' | 'resizeStart' | 'resizeTotals' | 'resizeOutside'

export interface TableChange {
  readonly area: TableArea
  readonly columnNames: readonly string[]
  readonly writes: readonly CellWrite[]
  readonly totals: WorkbookTableAdd['totals'] | undefined
}

/**
 * What a change of Header Row / Total Row does to the cells around a table, read from Excel 16
 * (`ListObject.ShowHeaders` / `ShowTotals` through COM):
 *  - Header Row off empties the header cells and the table starts one row lower (its column names
 *    stay with the table); on again, the header row is the row above the table (Excel inserts a row
 *    when that one is not free — here the change is refused instead, which leaves every other
 *    table and formula where it is);
 *  - Total Row on adds a row under the data: "Total" in the first column and `SUBTOTAL` in the last
 *    (a sum of numbers, a count of anything else); off again empties it. Excel makes room by taking
 *    the rows below into the table; here a row that is not empty refuses the change.
 */
export function planTableChange(input: {
  readonly area: TableArea
  readonly columnNames: readonly string[]
  readonly from: Pick<TableDesignOptions, 'headerRow' | 'totalRow'>
  readonly to: Pick<TableDesignOptions, 'headerRow' | 'totalRow'>
  readonly tableName: string
  readonly totalLabel: string
  readonly maxRows: number
  /** the value of a cell as it is now (null = empty) */
  readonly cellAt: (row: number, column: number) => unknown
}): TableChange | { readonly error: TableChangeError } {
  let area = input.area
  let names = [...input.columnNames]
  const writes: CellWrite[] = []
  let totals: WorkbookTableAdd['totals'] | undefined
  const columns = (): number[] => {
    const list: number[] = []
    for (let column = area.startColumn; column <= area.endColumn; column += 1) list.push(column)
    return list
  }
  const empty = (value: unknown): boolean => value === null || value === undefined || value === ''
  const free = (row: number): boolean =>
    columns().every((column) => empty(input.cellAt(row, column)))
  let header = input.from.headerRow
  let totalRow = input.from.totalRow

  // a Total Row that goes first, so the data rows it leaves are the ones the header change sees
  if (totalRow && !input.to.totalRow) {
    if (area.endRow - area.startRow + 1 - (header ? 1 : 0) - 1 < 1) return { error: 'needsData' }
    for (const column of columns()) writes.push({ row: area.endRow, column, value: null })
    area = { ...area, endRow: area.endRow - 1 }
    totalRow = false
  }

  if (header && !input.to.headerRow) {
    if (area.endRow - area.startRow + 1 - 1 - (totalRow ? 1 : 0) < 1) return { error: 'needsData' }
    // the header cells are the names now (they may have been edited since the table was made)
    // (made unique and short enough the way Excel names columns: a column's name is at most 255
    // characters, and two columns never share one in any case)
    const taken = new Set<string>()
    names = columns().map((column, index) => {
      const text = String(input.cellAt(area.startRow, column) ?? '').trim()
      const base = (text === '' ? (names[index] ?? `Column${index + 1}`) : text).slice(0, 250)
      let name = base
      for (let suffix = 2; taken.has(name.toLowerCase()); suffix += 1) name = `${base}${suffix}`
      taken.add(name.toLowerCase())
      return name
    })
    for (const column of columns()) writes.push({ row: area.startRow, column, value: null })
    area = { ...area, startRow: area.startRow + 1 }
    header = false
  } else if (!header && input.to.headerRow) {
    const row = area.startRow - 1
    if (row < 0 || !free(row)) return { error: 'noRoomAbove' }
    columns().forEach((column, index) => {
      writes.push({ row, column, value: names[index] ?? `Column${index + 1}` })
    })
    area = { ...area, startRow: row }
    header = true
  }

  if (!totalRow && input.to.totalRow) {
    const row = area.endRow + 1
    if (row >= input.maxRows || !free(row)) return { error: 'noRoomBelow' }
    const dataFirst = area.startRow + (header ? 1 : 0)
    const width = area.endColumn - area.startColumn + 1
    const lastIndex = width - 1
    const kind = totalsKindOf(
      Array.from({ length: area.endRow - dataFirst + 1 }, (_, i) =>
        input.cellAt(dataFirst + i, area.endColumn),
      ),
    )
    const meta: NonNullable<WorkbookTableAdd['totals']> = names.map(() => null)
    if (width > 1) {
      writes.push({ row, column: area.startColumn, value: input.totalLabel })
      meta[0] = { label: input.totalLabel }
    }
    writes.push({
      row,
      column: area.endColumn,
      value: totalsFormula(kind, input.tableName, names[lastIndex] ?? `Column${width}`),
    })
    meta[lastIndex] = { ...(meta[lastIndex] ?? {}), fn: kind }
    totals = meta
    area = { ...area, endRow: row }
    totalRow = true
  }

  return { area, columnNames: names, writes, totals }
}

/** a typed range ("A1:D10", "$A$1:$D$10") as an area, or null when it does not read as one */
export function parseAreaText(text: string): TableArea | null {
  const match =
    /^\s*=?\s*\$?([A-Za-z]{1,3})\$?(\d+)(?:\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+))?\s*$/.exec(text)
  if (!match) return null
  const index = (label: string): number =>
    [...label.toUpperCase()].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1
  const startColumn = index(match[1]!)
  const startRow = Number(match[2]) - 1
  const endColumn = match[3] === undefined ? startColumn : index(match[3])
  const endRow = match[4] === undefined ? startRow : Number(match[4]) - 1
  if (startRow < 0 || endRow < 0) return null
  return {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startColumn: Math.min(startColumn, endColumn),
    endColumn: Math.max(startColumn, endColumn),
  }
}

/**
 * Resize Table (Excel: the header row stays in its row, and the new range overlaps the old one).
 * Here the table's top-left cell stays too, so the columns it already has keep their names.
 */
export function checkResize(
  area: TableArea,
  next: TableArea,
  options: Pick<TableDesignOptions, 'headerRow' | 'totalRow'>,
): TableChangeError | null {
  if (options.totalRow) return 'resizeTotals'
  // every column is a name, a header cell and a part of the table: a width nobody means is a typo
  if (next.endColumn - next.startColumn + 1 > MAX_TABLE_COLUMNS) return 'resizeOutside'
  if (next.startRow !== area.startRow || next.startColumn !== area.startColumn) return 'resizeStart'
  const dataRows = next.endRow - next.startRow + 1 - (options.headerRow ? 1 : 0)
  if (dataRows < 1) return 'needsData'
  return null
}
