import { describe, expect, it, vi } from 'vitest'
import { ICommandService, IUndoRedoService } from '@univerjs/core'
import {
  AddWorksheetMergeMutation,
  RemoveWorksheetMergeMutation,
  SelectionMoveType,
  SetSelectionsOperation,
  SheetsSelectionsService,
} from '@univerjs/sheets'

import { unmergePreservingCellStyles } from '../src/renderer/unmerge-preserve-styles'
import type { UniverRuntime, UniverWorksheet } from '../src/renderer/univer-state'

function setup(merges = [{ row: 0, column: 0, height: 2, width: 2 }]) {
  const execute = vi.fn((_id: string, _params: unknown) => true)
  const pushUndoRedo = vi.fn()
  const selections = [
    {
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      primary: { actualRow: 0, actualColumn: 0, isMerged: true, isMergedMainCell: true },
    },
  ]
  const sheet = {
    getSheetId: () => 'sheet-1',
    getMergeData: () =>
      merges.map((merge) => ({
        getRow: () => merge.row,
        getColumn: () => merge.column,
        getHeight: () => merge.height,
        getWidth: () => merge.width,
      })),
  } as unknown as UniverWorksheet
  const runtime = {
    univerAPI: { getActiveWorkbook: () => ({ getId: () => 'workbook-1' }) },
    univer: {
      __getInjector: () => ({
        get: (token: unknown) => {
          if (token === ICommandService) return { syncExecuteCommand: execute }
          if (token === IUndoRedoService) return { pushUndoRedo }
          if (token === SheetsSelectionsService) return { getCurrentSelections: () => selections }
          throw new Error('Unexpected service')
        },
      }),
    },
  } as unknown as UniverRuntime
  return { sheet, runtime, execute, pushUndoRedo, selections }
}

describe('unmergePreservingCellStyles', () => {
  it('removes only the actual merge and records an Undo/Redo pair without copying the anchor XF', () => {
    const { sheet, runtime, execute, pushUndoRedo, selections } = setup()
    expect(
      unmergePreservingCellStyles(runtime, sheet, {
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      }),
    ).toBe(true)
    const params = {
      unitId: 'workbook-1',
      subUnitId: 'sheet-1',
      ranges: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }],
    }
    expect(execute).toHaveBeenNthCalledWith(1, RemoveWorksheetMergeMutation.id, params)
    expect(execute).toHaveBeenNthCalledWith(
      2,
      SetSelectionsOperation.id,
      expect.objectContaining({
        type: SelectionMoveType.ONLY_SET,
        selections: [
          expect.objectContaining({ primary: expect.objectContaining({ isMerged: false }) }),
        ],
      }),
    )
    expect(pushUndoRedo).toHaveBeenCalledExactlyOnceWith(
      expect.objectContaining({
        unitID: 'workbook-1',
        undoMutations: [
          { id: AddWorksheetMergeMutation.id, params },
          { id: SetSelectionsOperation.id, params: expect.any(Object) },
        ],
        redoMutations: [
          { id: RemoveWorksheetMergeMutation.id, params },
          { id: SetSelectionsOperation.id, params: expect.any(Object) },
        ],
      }),
    )
    expect(selections[0]?.primary.isMerged).toBe(true)
  })

  it('unmerges intersecting actual ranges and leaves unrelated neighboring merges alone', () => {
    const { sheet, runtime, execute } = setup([
      { row: 0, column: 0, height: 2, width: 2 },
      { row: 0, column: 3, height: 2, width: 2 },
    ])
    expect(
      unmergePreservingCellStyles(runtime, sheet, {
        startRow: 1,
        endRow: 1,
        startColumn: 1,
        endColumn: 1,
      }),
    ).toBe(true)
    expect(execute.mock.calls[0]?.[1]).toMatchObject({
      ranges: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }],
    })
  })

  it('does not create an undo item when there is no merge or removal fails', () => {
    const noMerge = setup([])
    expect(
      unmergePreservingCellStyles(noMerge.runtime, noMerge.sheet, {
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 0,
      }),
    ).toBe(false)
    expect(noMerge.execute).not.toHaveBeenCalled()
    const failed = setup()
    failed.execute.mockReturnValue(false)
    expect(
      unmergePreservingCellStyles(failed.runtime, failed.sheet, {
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      }),
    ).toBe(false)
    expect(failed.pushUndoRedo).not.toHaveBeenCalled()
  })
})
