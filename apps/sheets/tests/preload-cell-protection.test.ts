import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  contextBridge: { exposeInMainWorld: vi.fn() },
  ipcRenderer: { invoke: vi.fn(), on: vi.fn(), removeListener: vi.fn() },
  webUtils: { getPathForFile: vi.fn() },
}))
vi.mock('@genoffice/electron-utils/drop-open', () => ({ installDropOpenBridge: vi.fn() }))

import { parseCellStyle } from '../src/preload/index'
import { createEditJournal } from '../src/renderer/edit-journal'
import { selectionProtection } from '../src/renderer/selection-protection'
import type { WorkbookFile } from '../src/shared/desktop-api'

const base = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

describe('native style preload protection boundary', () => {
  it('retains native optional flags through preload to selected-cell backfill', () => {
    const styles = [
      parseCellStyle(base),
      parseCellStyle({ ...base, protectionLocked: false }),
      parseCellStyle({ ...base, protectionHidden: true }),
    ]
    const selected = selectionProtection(
      styles,
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      {
        cells: [
          { row: 0, column: 0, value: 'A', styleIndex: 1 },
          { row: 0, column: 1, value: 5, formula: '=2+3', styleIndex: 2 },
          { row: 0, column: 2, value: null, styleIndex: 2 },
        ],
        rows: [],
      },
      createEditJournal(),
      'sheet-1',
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
    )
    expect(styles[1]).toMatchObject({ protectionLocked: false })
    expect(styles[2]).toMatchObject({ protectionHidden: true })
    expect(selected.cells.map(({ locked, hidden }) => [locked, hidden])).toEqual([
      [false, false],
      [true, true],
      [true, true],
    ])
  })

  it('rejects malformed protection flags instead of treating them as OOXML defaults', () => {
    expect(() => parseCellStyle({ ...base, protectionLocked: 'false' })).toThrow(
      'Invalid workbook style response.',
    )
  })

  it('retains native accounting underline and script position for selected-cell backfill', () => {
    const styles = [
      parseCellStyle(base),
      parseCellStyle({
        ...base,
        underline: true,
        underlineStyle: 'singleAccounting',
        fontVerticalAlign: 'superscript',
      }),
      parseCellStyle({
        ...base,
        underline: true,
        underlineStyle: 'doubleAccounting',
        fontVerticalAlign: 'subscript',
      }),
    ]
    const read = {
      cells: [
        { row: 0, column: 0, value: 'A', styleIndex: 1 },
        { row: 0, column: 1, value: 5, styleIndex: 2 },
      ],
      rows: [],
    }
    const sheet = { columnWidths: [] } as unknown as WorkbookFile['sheets'][number]
    const journal = createEditJournal()
    expect(
      selectionProtection(styles, sheet, read, journal, 'sheet-1', {
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 0,
      }).font,
    ).toEqual({ underlineStyle: 'singleAccounting', verticalAlign: 'superscript' })
    expect(
      selectionProtection(styles, sheet, read, journal, 'sheet-1', {
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 1,
      }).font,
    ).toEqual({ underlineStyle: null, verticalAlign: null })
    expect(() => parseCellStyle({ ...base, fontVerticalAlign: 'raised' })).toThrow(
      'Invalid workbook style response.',
    )
  })
})
