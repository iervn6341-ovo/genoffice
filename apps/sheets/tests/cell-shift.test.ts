import { describe, expect, it } from 'vitest'
import {
  CELL_SHIFT_COMMAND_IDS,
  cellShiftAction,
  cellShiftOptions,
  defaultCellShift,
  directCellShift,
  parseCellShiftCommand,
} from '../src/renderer/cell-shift'

// C4: Excel's Insert Cells / Delete Cells dialog → command mapping
const B2_C4 = { startRow: 1, endRow: 3, startColumn: 1, endColumn: 2 } // 3 rows × 2 cols
const B2_D2 = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 3 } // 1 row × 3 cols

describe('cell shift mapping (C4)', () => {
  it('lists Excel options in order', () => {
    expect(cellShiftOptions('insert')).toEqual(['shift-right', 'shift-down', 'row', 'column'])
    expect(cellShiftOptions('delete')).toEqual(['shift-left', 'shift-up', 'row', 'column'])
  })

  it('preselects by selection shape like Excel', () => {
    expect(defaultCellShift('insert', B2_C4)).toBe('shift-right')
    expect(defaultCellShift('insert', B2_D2)).toBe('shift-down')
    expect(defaultCellShift('delete', B2_C4)).toBe('shift-left')
    expect(defaultCellShift('delete', B2_D2)).toBe('shift-up')
    // a single cell counts as "not wider": shift right / left
    const one = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
    expect(defaultCellShift('insert', one)).toBe('shift-right')
  })

  it('maps range shifts to the Univer commands', () => {
    expect(cellShiftAction('insert', 'shift-right', B2_C4)).toEqual({
      kind: 'command',
      id: 'sheet.command.insert-range-move-right',
      range: B2_C4,
    })
    expect(cellShiftAction('insert', 'shift-down', B2_C4)).toMatchObject({
      id: 'sheet.command.insert-range-move-down',
    })
    expect(cellShiftAction('delete', 'shift-left', B2_C4)).toMatchObject({
      id: 'sheet.command.delete-range-move-left',
    })
    expect(cellShiftAction('delete', 'shift-up', B2_C4)).toMatchObject({
      id: 'sheet.command.delete-range-move-up',
    })
    for (const id of [
      'sheet.command.insert-range-move-right',
      'sheet.command.insert-range-move-down',
      'sheet.command.delete-range-move-left',
      'sheet.command.delete-range-move-up',
    ]) {
      expect(CELL_SHIFT_COMMAND_IDS.has(id)).toBe(true)
    }
  })

  it('rejects a shift direction from the other dialog', () => {
    expect(cellShiftAction('insert', 'shift-up', B2_C4)).toBeNull()
    expect(cellShiftAction('delete', 'shift-right', B2_C4)).toBeNull()
  })

  it('entire row / column cover the selection span', () => {
    expect(cellShiftAction('insert', 'row', B2_C4)).toEqual({
      kind: 'rows',
      op: 'insert_rows',
      row: 2,
      count: 3,
    })
    expect(cellShiftAction('delete', 'row', B2_C4)).toMatchObject({ op: 'delete_rows', count: 3 })
    expect(cellShiftAction('insert', 'column', B2_C4)).toEqual({
      kind: 'columns',
      op: 'insert_cols',
      column: 1,
      count: 2,
    })
    expect(cellShiftAction('delete', 'column', B2_C4)).toMatchObject({ op: 'delete_cols' })
  })

  it('whole rows / columns skip the dialog', () => {
    expect(
      directCellShift({ startRow: 4, endRow: 5, startColumn: 0, endColumn: 99 }, 1000, 100),
    ).toBe('row')
    expect(
      directCellShift({ startRow: 0, endRow: 999, startColumn: 2, endColumn: 2 }, 1000, 100),
    ).toBe('column')
    expect(directCellShift(B2_C4, 1000, 100)).toBeNull()
  })

  it('parses the ribbon command string', () => {
    expect(parseCellShiftCommand('cells-shift:insert:shift-down')).toEqual({
      mode: 'insert',
      option: 'shift-down',
    })
    expect(parseCellShiftCommand('cells-shift:delete:column')).toEqual({
      mode: 'delete',
      option: 'column',
    })
    expect(parseCellShiftCommand('cells-shift:insert:sideways')).toBeNull()
    expect(parseCellShiftCommand('insert-row-here')).toBeNull()
  })
})
