import { describe, expect, it, vi } from 'vitest'

import { captureSymbolCellTarget, handleInsertSymbol } from '../src/renderer/data-tools-actions'

describe('plain Symbol destination after editor Enter', () => {
  it('appends to the cell that was editing before Enter moved the active range', () => {
    let activeRow = 0
    let workbookId = 'book-1'
    const original = {
      getFormula: () => '',
      getCellDatas: () => [[{ v: 'Draft' }]],
      getValue: () => 'Draft',
      setValue: vi.fn(),
    }
    const nextRow = {
      getFormula: () => '',
      getCellDatas: () => [[{ v: null }]],
      getValue: () => null,
      setValue: vi.fn(),
    }
    const sheet = {
      getSheetId: () => 'sheet-1',
      getSheet: () => ({ getComposedCellStyle: () => ({}) }),
      getRange: (row: number, column: number) => {
        expect(column).toBe(2)
        return row === 0 ? original : nextRow
      },
    }
    const workbook = {
      getId: () => workbookId,
      getActiveSheet: () => sheet,
      getSheetBySheetId: (id: string) => (id === 'sheet-1' ? sheet : null),
      getActiveRange: () => ({ getRow: () => activeRow, getColumn: () => 2 }),
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } } as never
    const message = vi.fn()
    const ctx = {
      univerRef: { current: runtime },
      lazyWorkbookRef: { current: null },
      setMessage: message,
    } as never

    const target = captureSymbolCellTarget(runtime)
    expect(target).toEqual({ workbookId: 'book-1', sheetId: 'sheet-1', row: 0, column: 2 })
    activeRow = 1 // FWorkbook.endEditingAsync(true) sends Enter and advances C1 → C2.
    handleInsertSymbol(ctx, 'Ω', 'Cambria Math', target)
    expect((original.setValue.mock.calls[0]?.[0] as { v: string }).v).toBe('DraftΩ')
    expect(nextRow.setValue).not.toHaveBeenCalled()
    expect(message).toHaveBeenCalledWith(expect.stringContaining('C1'))

    workbookId = 'another-book'
    handleInsertSymbol(ctx, '∑', 'Cambria Math', target)
    expect(original.setValue).toHaveBeenCalledTimes(1)
    expect(nextRow.setValue).not.toHaveBeenCalled()
  })
})
