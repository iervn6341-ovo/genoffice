/**
 * Excel-style horizontal "infinite" columns.
 *
 * Excel keeps every column up to XFD (16,384) available. Materialising all
 * of them up front would make every whole-row operation, select-all, formula
 * range scan and skeleton rebuild 600x wider, so the grid instead grows on
 * demand: when the viewport, the selection, a jump (Ctrl+Arrow), Go To or the
 * Name Box gets close to the last column, the sheet's column count is raised
 * in a batch, capped at Excel's limit.
 *
 * The growth runs the bare `SetWorksheetColumnCountMutation` through the
 * command service. Unlike the facade's `setColumnCount` (a command) it pushes
 * no undo step, and the edit journal does not list the mutation, so growth
 * never dirties the workbook, never reaches a save and never changes the
 * file's `<dimension>` / `<cols>` — only cells the user later fills do.
 */
import { ICommandService, RANGE_TYPE } from '@univerjs/core'
import type { IStyleData, Worksheet } from '@univerjs/core'
import { SetWorksheetColumnCountMutation } from '@univerjs/sheets'
import type { UniverRuntime } from './univer-state'

/** Excel's last column, XFD, as a column count. */
export const EXCEL_MAX_COLUMNS = 16384

/** Always keep at least this many spare columns right of the reach. */
const MIN_SPARE_COLUMNS = 26

/**
 * The column count to grow to, or null when the grid is already roomy enough.
 *
 * @param count current column count of the sheet
 * @param reachColumn the rightmost 0-based column that must stay reachable
 * @param visibleSpan columns visible in the viewport (spare room scales with it)
 */
export function planColumnExtension(
  count: number,
  reachColumn: number,
  visibleSpan = 0,
  maxColumns = EXCEL_MAX_COLUMNS,
): number | null {
  if (!Number.isFinite(reachColumn) || reachColumn < 0) return null
  if (count >= maxColumns) return null
  const needed = Math.floor(reachColumn) + 1
  const spare = Math.max(MIN_SPARE_COLUMNS, Math.ceil(visibleSpan) + 1)
  if (count - needed >= spare) return null
  const target = Math.min(maxColumns, needed + spare * 2)
  return target > count ? target : null
}

/** Column count that exactly reaches `columnIndex` (explicit jump), or null. */
export function columnCountToReach(
  count: number,
  columnIndex: number,
  maxColumns = EXCEL_MAX_COLUMNS,
): number | null {
  if (!Number.isFinite(columnIndex) || columnIndex < 0) return null
  const needed = Math.min(maxColumns, Math.floor(columnIndex) + 1)
  return needed > count ? needed : null
}

type ColumnData = Record<number, { w?: number; hd?: number; s?: IStyleData }>

/**
 * Supplies the file's `<col>` widths / styles / hidden flags for newly added
 * columns (a workbook-wide `<col max="16384">` must keep applying to them).
 */
export type ColumnDataSource = (
  sheetId: string,
  fromColumn: number,
  toColumnExclusive: number,
) => ColumnData | null

let columnDataSource: ColumnDataSource | null = null

export function setColumnDataSource(source: ColumnDataSource | null): void {
  columnDataSource = source
}

/**
 * Grows `worksheet` to `newCount` columns without touching undo history.
 * Returns true when the column count changed.
 */
export function extendColumns(
  commandService: ICommandService,
  unitId: string,
  worksheet: Worksheet,
  newCount: number,
): boolean {
  const count = Math.min(Math.floor(newCount), EXCEL_MAX_COLUMNS)
  const current = worksheet.getColumnCount()
  if (count <= current) return false
  const sheetId = worksheet.getSheetId()
  const fill = columnDataSource?.(sheetId, current, count)
  if (fill) {
    const columnData = worksheet.getConfig().columnData as ColumnData
    for (const [key, value] of Object.entries(fill)) {
      const column = Number(key)
      columnData[column] = { ...columnData[column], ...value }
    }
  }
  return commandService.syncExecuteCommand(SetWorksheetColumnCountMutation.id, {
    unitId,
    subUnitId: sheetId,
    columnCount: count,
  })
}

/** Grows the sheet so `columnIndex` exists. Returns true when it grew. */
export function ensureColumnReachable(
  commandService: ICommandService,
  unitId: string,
  worksheet: Worksheet,
  columnIndex: number,
): boolean {
  const target = columnCountToReach(worksheet.getColumnCount(), columnIndex)
  return target === null ? false : extendColumns(commandService, unitId, worksheet, target)
}

/** Facade-level variant for UI code that holds a `univerAPI` worksheet. */
export function ensureColumnOnActiveWorkbook(
  runtime: UniverRuntime,
  sheetId: string | null,
  columnIndex: number,
): boolean {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return false
  const sheet = sheetId ? workbook.getSheetBySheetId(sheetId) : workbook.getActiveSheet()
  if (!sheet) return false
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  return ensureColumnReachable(commandService, workbook.getId(), sheet.getSheet(), columnIndex)
}

/**
 * Keeps spare columns to the right of the viewport and of the selection.
 * Whole-row / select-all selections are ignored: their right edge is the
 * grid edge, not a place the user is working at.
 */
export function installColumnAutoExtend(runtime: UniverRuntime): () => void {
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  const grow = (reach: number, span: number): void => {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    if (!workbook || !sheet) return
    const core = sheet.getSheet()
    const target = planColumnExtension(core.getColumnCount(), reach, span)
    if (target !== null) extendColumns(commandService, workbook.getId(), core, target)
  }
  const onScroll = runtime.univerAPI.addEvent(runtime.univerAPI.Event.Scroll, (params) => {
    try {
      const visible = params.worksheet.getVisibleRange()
      if (visible) grow(visible.endColumn, visible.endColumn - visible.startColumn)
    } catch {
      // no scroll render controller yet (still booting)
    }
  })
  const onSelection = runtime.univerAPI.addEvent(runtime.univerAPI.Event.SelectionChanged, () => {
    try {
      const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      const ranges = sheet?.getSelection()?.getActiveRangeList() ?? []
      let reach = -1
      for (const range of ranges) {
        const bounds = range.getRange()
        if (bounds.rangeType === RANGE_TYPE.ROW || bounds.rangeType === RANGE_TYPE.ALL) continue
        reach = Math.max(reach, bounds.endColumn)
      }
      if (reach >= 0) grow(reach, 0)
    } catch {
      // selection raced a workbook swap
    }
  })
  return () => {
    onScroll.dispose()
    onSelection.dispose()
  }
}
