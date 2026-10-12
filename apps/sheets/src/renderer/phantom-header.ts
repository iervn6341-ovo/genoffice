/**
 * A table without a header row, in Univer's model, is a table that starts one row higher — over
 * the empty row its header left — because Univer counts structured references (Table1[Col]) from
 * the row below the table's first and has no table without a header. Univer also writes the
 * column names into the first row of every table it holds, wherever the cell is empty; this keeps
 * that row empty on screen (it is empty in Excel), without writing anything into the cells.
 */
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

import type { WorkbookTableAdd } from '../shared/desktop-api'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/** true when (row, column) of the sheet is a cell of the empty row above a table without a header */
export function isPhantomHeaderCell(
  tables: readonly WorkbookTableAdd[],
  sheetId: string,
  row: number,
  column: number,
): boolean {
  return tables.some(
    (table) =>
      table.headerRow === false &&
      table.sheetId === sheetId &&
      row === table.area.startRow - 1 &&
      column >= table.area.startColumn &&
      column <= table.area.endColumn,
  )
}

export function installPhantomHeaderInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // above Univer's own table interceptor (priority 0), which looks for a cell with no value
    priority: 100_000,
    handler: (cell, location, next) => {
      if (cell?.v !== undefined) return next(cell)
      const tables = lazyWorkbookRef.current?.editJournal.tableAdds
      if (!tables || tables.length === 0) return next(cell)
      if (!isPhantomHeaderCell(tables, location.subUnitId, location.row, location.col)) {
        return next(cell)
      }
      return next({ ...(cell ?? location.rawData ?? {}), v: '' })
    },
  })
}
