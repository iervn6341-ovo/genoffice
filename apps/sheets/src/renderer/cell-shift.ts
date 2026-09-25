/// Excel's Insert Cells (⌃⇧= / ⌘⇧+) and Delete Cells (⌃- / ⌘-) dialogs: the
/// four options map onto Univer's range-shift commands or onto whole row /
/// column structural ops. Pure so the mapping is unit-testable.

export type CellShiftMode = 'insert' | 'delete'
export type CellShiftOption =
  'shift-right' | 'shift-down' | 'shift-left' | 'shift-up' | 'row' | 'column'

export interface CellShiftRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/** the dialog's four radios, in Excel's order */
export function cellShiftOptions(mode: CellShiftMode): readonly CellShiftOption[] {
  return mode === 'insert'
    ? ['shift-right', 'shift-down', 'row', 'column']
    : ['shift-left', 'shift-up', 'row', 'column']
}

/**
 * Excel preselects by the selection's shape: a selection wider than it is tall
 * shifts vertically (insert down / delete up), otherwise horizontally.
 */
export function defaultCellShift(mode: CellShiftMode, range: CellShiftRange): CellShiftOption {
  const rows = range.endRow - range.startRow + 1
  const columns = range.endColumn - range.startColumn + 1
  const vertical = columns > rows
  if (mode === 'insert') return vertical ? 'shift-down' : 'shift-right'
  return vertical ? 'shift-up' : 'shift-left'
}

export const CELL_SHIFT_COMMANDS = {
  'shift-right': 'sheet.command.insert-range-move-right',
  'shift-down': 'sheet.command.insert-range-move-down',
  'shift-left': 'sheet.command.delete-range-move-left',
  'shift-up': 'sheet.command.delete-range-move-up',
} as const

/** the range-shift command ids — they rewrite cell content like move-range, so they share its gating */
export const CELL_SHIFT_COMMAND_IDS: ReadonlySet<string> = new Set(
  Object.values(CELL_SHIFT_COMMANDS),
)

export type CellShiftAction =
  | { readonly kind: 'command'; readonly id: string; readonly range: CellShiftRange }
  | {
      readonly kind: 'rows'
      readonly op: 'insert_rows' | 'delete_rows'
      /** 1-based first row */
      readonly row: number
      readonly count: number
    }
  | {
      readonly kind: 'columns'
      readonly op: 'insert_cols' | 'delete_cols'
      /** 0-based first column */
      readonly column: number
      readonly count: number
    }

/** what one dialog choice does to the selection */
export function cellShiftAction(
  mode: CellShiftMode,
  option: CellShiftOption,
  range: CellShiftRange,
): CellShiftAction | null {
  if (option === 'row') {
    return {
      kind: 'rows',
      op: mode === 'insert' ? 'insert_rows' : 'delete_rows',
      row: range.startRow + 1,
      count: range.endRow - range.startRow + 1,
    }
  }
  if (option === 'column') {
    return {
      kind: 'columns',
      op: mode === 'insert' ? 'insert_cols' : 'delete_cols',
      column: range.startColumn,
      count: range.endColumn - range.startColumn + 1,
    }
  }
  const allowed = cellShiftOptions(mode)
  if (!allowed.includes(option)) return null
  return { kind: 'command', id: CELL_SHIFT_COMMANDS[option], range }
}

/**
 * Excel skips the dialog when whole rows or whole columns are selected and
 * inserts/deletes them directly.
 */
export function directCellShift(
  range: CellShiftRange,
  rowCount: number,
  columnCount: number,
): CellShiftOption | null {
  if (range.startColumn === 0 && range.endColumn >= columnCount - 1) return 'row'
  if (range.startRow === 0 && range.endRow >= rowCount - 1) return 'column'
  return null
}

/** parses the `cells-shift:<mode>:<option>` command string */
export function parseCellShiftCommand(
  command: string,
): { mode: CellShiftMode; option: CellShiftOption } | null {
  const match =
    /^cells-shift:(insert|delete):(shift-right|shift-down|shift-left|shift-up|row|column)$/.exec(
      command,
    )
  if (!match) return null
  return { mode: match[1] as CellShiftMode, option: match[2] as CellShiftOption }
}
