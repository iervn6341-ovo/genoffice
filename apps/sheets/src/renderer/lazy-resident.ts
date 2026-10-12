import type { IRange } from '@univerjs/core'

import { ensureLazyRangeLoaded } from './univer-sync'
import { lazySheetScreenExtent, type LazyWorkbookState, type UniverRuntime } from './univer-state'

/**
 * Makes the part of `range` that the opened file has rows for resident in the grid before a bulk
 * edit reads or writes it. Cells past the file's own extent do not exist in the file and are
 * empty, so they need no load (a fresh workbook's file has a single row). True when the whole
 * range can be read from the model; false only when the file's part could not be loaded.
 */
export async function ensureRangeResident(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: Parameters<typeof ensureLazyRangeLoaded>[2],
  range: IRange,
  setMessage: (message: string) => void,
): Promise<boolean> {
  const state = lazyWorkbookRef.current
  if (!state) return true
  const extent = lazySheetScreenExtent(state, worksheet.getSheetId())
  if (!extent) return true
  const inside: IRange = {
    ...range,
    endRow: Math.min(range.endRow, extent.rows - 1),
    endColumn: Math.min(range.endColumn, extent.columns - 1),
  }
  if (inside.startRow > inside.endRow || inside.startColumn > inside.endColumn) return true
  return ensureLazyRangeLoaded(runtime, lazyWorkbookRef, worksheet, inside, setMessage)
}
