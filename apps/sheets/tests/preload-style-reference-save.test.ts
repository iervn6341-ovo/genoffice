import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  contextBridge: { exposeInMainWorld: vi.fn() },
  ipcRenderer: { invoke: vi.fn(), on: vi.fn(), removeListener: vi.fn() },
  webUtils: { getPathForFile: vi.fn() },
}))
vi.mock('@genoffice/electron-utils/drop-open', () => ({ installDropOpenBridge: vi.fn() }))

import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { parseSaveRequest } from '../src/preload/index'
import { createEditJournal, toSaveEdits } from '../src/renderer/edit-journal'
import { workbookSaveRequestSchema } from '../src/shared/desktop-api'
import { buildEditFixture } from './fixture-builder'

function saveRequest(edits: ReturnType<typeof toSaveEdits>) {
  return workbookSaveRequestSchema.parse({
    sessionId: '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a',
    mode: 'save',
    edits,
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
  })
}

describe('cross-save style reference through the preload save boundary', () => {
  it('accepts an Undo journal containing only original XF references and restores them in XLSX', async () => {
    const original = await buildEditFixture()
    const changed = await applyCellEditsToXlsx(original, [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { verticalAlignment: 'distributed' },
      },
    ])
    const journal = createEditJournal()
    journal.cells.set(
      'sheet-1',
      new Map([
        ['0:0', { row: 0, column: 0, hasValue: false, value: null, styleReference: 1 }],
        ['0:2', { row: 0, column: 2, hasValue: false, value: null, styleReference: null }],
      ]),
    )
    const parsed = parseSaveRequest(saveRequest(toSaveEdits(journal)))
    expect(parsed.edits.map((edit) => edit.styleReference)).toEqual([1, null])
    const restored = await applyCellEditsToXlsx(
      changed.buffer,
      parsed.edits.map((edit) => ({
        sheetName: 'Data',
        row: edit.row,
        column: edit.column,
        writeValue: edit.writeValue,
        cell: { value: edit.value, formula: edit.formula },
        styleReference: edit.styleReference,
      })),
    )
    const sheet = await (
      await JSZip.loadAsync(restored.buffer)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    expect(sheet).toContain('<c r="A1" t="s" s="1"><v>0</v></c>')
    expect(sheet).toContain('<c r="C1"><v>5</v></c>')
    expect(sheet).toContain('<f>SUM(C1:C2)</f><v>5</v>')
  })

  it('rejects malformed and conflicting style references at preload', () => {
    const entry = {
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
      writeValue: false,
      value: null,
      styleReference: 1,
    }
    const valid = saveRequest([entry])
    for (const invalid of [-1, 1.5, '1']) {
      expect(() =>
        parseSaveRequest({ ...valid, edits: [{ ...entry, styleReference: invalid }] } as never),
      ).toThrow('Invalid workbook cell edit.')
    }
    expect(() => parseSaveRequest({ ...valid, edits: [{ ...entry, styleReset: true }] })).toThrow(
      'Invalid workbook cell edit.',
    )
  })
})
