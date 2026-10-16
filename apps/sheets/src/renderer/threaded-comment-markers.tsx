/// The purple corner marker Excel 365 draws on a cell that holds a comment thread (notes keep
/// Univer's own red marker). One float-DOM layer per thread, anchored to its cell so it scrolls and
/// zooms with the grid; resolved threads are drawn muted.

import type { WorkbookSheetThread } from '../shared/desktop-api'
import type { createUniver } from './create-univer'

type UniverRuntime = ReturnType<typeof createUniver>

export interface ThreadMarkerHandle {
  /// Re-installs the markers when the thread set / active sheet changed since the last call.
  refresh(): void
  dispose(): void
}

const MAX_MARKERS = 500

export function installThreadMarkers(
  runtime: UniverRuntime,
  threadsOf: (sheetId: string) => readonly WorkbookSheetThread[],
): ThreadMarkerHandle {
  let disposables: { dispose(): void }[] = []
  let signature = ''
  const clear = (): void => {
    for (const item of disposables) item.dispose()
    disposables = []
  }
  return {
    refresh: () => {
      const facade = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (!facade) return
      const sheetId = facade.getSheetId()
      const threads = threadsOf(sheetId).slice(0, MAX_MARKERS)
      const next = `${sheetId}|${threads.map((t) => `${t.row}:${t.column}:${t.done ? 1 : 0}`).join(',')}`
      if (next === signature) return
      signature = next
      clear()
      for (const thread of threads) {
        const key = `thread-marker-${sheetId}-${thread.row}-${thread.column}-${thread.done ? 'r' : 'o'}`
        const width = facade.getColumnWidth(thread.column)
        const height = facade.getRowHeight(thread.row)
        disposables.push(
          runtime.univerAPI.registerComponent(key, () => (
            <svg
              className={`thread-marker${thread.done ? ' resolved' : ''}`}
              viewBox={`0 0 ${width} ${height}`}
              width={width}
              height={height}
              aria-hidden="true"
            >
              <polygon points={`${width - 9},0 ${width},0 ${width},9`} />
            </svg>
          )),
        )
        const floating = facade.addFloatDomToRange(
          facade.getRange(thread.row, thread.column, 1, 1),
          { componentKey: key, allowTransform: false, eventPassThrough: true },
          {},
          key,
        )
        if (floating) disposables.push(floating)
      }
    },
    dispose: () => {
      signature = ''
      clear()
    },
  }
}
