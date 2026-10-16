import { ICommandService, IUndoRedoService, type IRange } from '@univerjs/core'
import {
  AddWorksheetMergeMutation,
  RemoveWorksheetMergeMutation,
  SelectionMoveType,
  SetSelectionsOperation,
  SheetsSelectionsService,
} from '@univerjs/sheets'

import type { UniverRuntime, UniverWorksheet } from './univer-state'

type MergeBounds = Pick<IRange, 'startRow' | 'endRow' | 'startColumn' | 'endColumn'>

function intersects(left: MergeBounds, right: MergeBounds): boolean {
  return (
    left.startRow <= right.endRow &&
    left.endRow >= right.startRow &&
    left.startColumn <= right.endColumn &&
    left.endColumn >= right.startColumn
  )
}

/**
 * Excel leaves each covered cell's XF untouched when a merge is removed.
 * Univer's breakApart command copies the anchor's entire style to every
 * covered cell, so use its merge-only mutation and put that mutation on the
 * ordinary Undo/Redo stack without the style-copy mutation.
 */
export function unmergePreservingCellStyles(
  runtime: UniverRuntime,
  sheet: UniverWorksheet,
  selection: MergeBounds,
): boolean {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return false
  const ranges: IRange[] = sheet
    .getMergeData()
    .map((merged) => ({
      startRow: merged.getRow(),
      endRow: merged.getRow() + merged.getHeight() - 1,
      startColumn: merged.getColumn(),
      endColumn: merged.getColumn() + merged.getWidth() - 1,
    }))
    .filter((merged) => intersects(merged, selection))
  if (ranges.length === 0) return false

  const params = { unitId: workbook.getId(), subUnitId: sheet.getSheetId(), ranges }
  const injector = runtime.univer.__getInjector()
  const commands = injector.get(ICommandService)
  const before = injector
    .get(SheetsSelectionsService)
    .getCurrentSelections()
    .map((item) => ({
      ...item,
      range: { ...item.range },
      primary: item.primary ? { ...item.primary } : item.primary,
    }))
  const after = before.map((item, index) => {
    if (index !== before.length - 1) return item
    const { startRow, startColumn } = item.range
    return {
      ...item,
      primary: {
        startRow,
        startColumn,
        endRow: startRow,
        endColumn: startColumn,
        actualRow: startRow,
        actualColumn: startColumn,
        isMerged: false,
        isMergedMainCell: false,
      },
    }
  })
  const selectionParams = (selections: typeof before) => ({
    unitId: params.unitId,
    subUnitId: params.subUnitId,
    selections,
    type: SelectionMoveType.ONLY_SET,
  })
  if (!commands.syncExecuteCommand(RemoveWorksheetMergeMutation.id, params)) return false
  if (
    before.length > 0 &&
    !commands.syncExecuteCommand(SetSelectionsOperation.id, selectionParams(after))
  ) {
    commands.syncExecuteCommand(AddWorksheetMergeMutation.id, params)
    return false
  }
  injector.get(IUndoRedoService).pushUndoRedo({
    unitID: params.unitId,
    undoMutations: [
      { id: AddWorksheetMergeMutation.id, params },
      ...(before.length > 0
        ? [{ id: SetSelectionsOperation.id, params: selectionParams(before) }]
        : []),
    ],
    redoMutations: [
      { id: RemoveWorksheetMergeMutation.id, params },
      ...(before.length > 0
        ? [{ id: SetSelectionsOperation.id, params: selectionParams(after) }]
        : []),
    ],
  })
  return true
}
