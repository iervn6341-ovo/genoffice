/**
 * Home → Fill: Up and Left (the Univer fill command run the other way — Excel copies the bottom row
 * up and the right column left, or the neighbor below / right of a single cell) and Series…
 * (see fill-series.ts for the planning rules).
 */
import { CellValueType, numfmt, type ICellData, type IRange, type IStyleData } from '@univerjs/core'

import { AUTO_FILL_COMMAND } from './app-constants'
import type { FillSeriesSeed } from './FillSeriesDialog'
import { t, type StringKey } from './i18n/locale'
import {
  defaultSeriesAxis,
  planSeries,
  type SeriesAxis,
  type SeriesDateUnit,
  type SeriesOptions,
  type SeriesRect,
  type SeriesSkip,
  type SeriesType,
} from './fill-series'
import type { RibbonCommandContext } from './ribbon-actions'
import { runAsOneUndoStepAsync } from './undo-group'
import { ensureRangeResident } from './lazy-resident'
import type { UniverRuntime } from './univer-state'

/// Univer's applyType values (AUTO_FILL_APPLY_TYPE is not exported at runtime)
const APPLY_COPY = 'COPY'
const APPLY_SERIES = 'SERIES'

/// What the Series dialog hands back. `autofill` is Excel's AutoFill type: the drag-fill rules.
export interface FillSeriesParams {
  readonly axis: SeriesAxis
  readonly type: SeriesType | 'autofill'
  readonly dateUnit: SeriesDateUnit
  readonly step: number
  readonly stop: number | null
  readonly trend: boolean
}

interface Selection {
  readonly unitId: string
  readonly sheetId: string
  readonly range: IRange
}

function currentSelection(runtime: UniverRuntime): Selection | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !sheet || !active) return null
  return {
    unitId: workbook.getId(),
    sheetId: sheet.getSheetId(),
    range: {
      startRow: active.getRow(),
      endRow: active.getRow() + active.getHeight() - 1,
      startColumn: active.getColumn(),
      endColumn: active.getColumn() + active.getWidth() - 1,
    },
  }
}

/// The source / target pair for Fill Up and Fill Left. A one-row (one-column) selection copies the
/// row below (the column to the right) like Excel; a bigger one copies its own bottom row (right
/// column). Null when there is no neighbor to copy from (the sheet's last row / column).
export function fillTowardsRanges(
  direction: 'up' | 'left',
  selection: IRange,
  maxRows: number,
  maxColumns: number,
): { sourceRange: IRange; targetRange: IRange } | null {
  const { startRow, endRow, startColumn, endColumn } = selection
  if (direction === 'up') {
    if (startRow === endRow) {
      if (endRow + 1 >= maxRows) return null
      return {
        sourceRange: { startRow: endRow + 1, endRow: endRow + 1, startColumn, endColumn },
        targetRange: { startRow, endRow: endRow + 1, startColumn, endColumn },
      }
    }
    return {
      sourceRange: { startRow: endRow, endRow, startColumn, endColumn },
      targetRange: { startRow, endRow, startColumn, endColumn },
    }
  }
  if (startColumn === endColumn) {
    if (endColumn + 1 >= maxColumns) return null
    return {
      sourceRange: { startRow, endRow, startColumn: endColumn + 1, endColumn: endColumn + 1 },
      targetRange: { startRow, endRow, startColumn, endColumn: endColumn + 1 },
    }
  }
  return {
    sourceRange: { startRow, endRow, startColumn: endColumn, endColumn },
    targetRange: { startRow, endRow, startColumn, endColumn },
  }
}

export function fillTowards(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  direction: 'up' | 'left',
): void {
  const selection = currentSelection(runtime)
  const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!selection || !sheet) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const ranges = fillTowardsRanges(
    direction,
    selection.range,
    sheet.getMaxRows(),
    sheet.getMaxColumns(),
  )
  if (!ranges) return
  void runtime.univerAPI.executeCommand(AUTO_FILL_COMMAND, {
    ...ranges,
    unitId: selection.unitId,
    subUnitId: selection.sheetId,
    applyType: APPLY_COPY,
  })
}

/// What the Series dialog opens on: along the row for a one-row selection, a Date type when the
/// first cell carries a date format.
export function fillSeriesSeed(runtime: UniverRuntime): FillSeriesSeed | null {
  const active = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!active) return null
  let startIsDate = false
  try {
    const pattern = active.getNumberFormat()
    startIsDate = pattern !== '' && numfmt.isDateFormat(pattern)
  } catch {
    // an unformatted cell has no pattern to read
  }
  return {
    axis: defaultSeriesAxis({
      startRow: active.getRow(),
      startColumn: active.getColumn(),
      rowCount: active.getHeight(),
      columnCount: active.getWidth(),
    }),
    startIsDate,
  }
}

const SKIP_MESSAGES: Record<SeriesSkip, StringKey> = {
  'no-numeric-start': 'appSeriesNoStart',
  'stop-required': 'appSeriesNeedsStop',
  'too-large': 'appSeriesTooLarge',
  'trend-needs-positive': 'appSeriesTrendPositive',
  'invalid-number': 'appSeriesInvalidNumber',
}

function cellNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

/// AutoFill type: Univer's drag-fill with the leading filled cells of the first line as the seed
async function autofillSeries(
  runtime: UniverRuntime,
  selection: Selection,
  axis: SeriesAxis,
  values: readonly (readonly unknown[])[],
): Promise<void> {
  const { range } = selection
  const filled = (value: unknown): boolean => value !== null && value !== undefined && value !== ''
  if (axis === 'columns') {
    const rows = values.length
    let seed = rows
    for (let c = 0; c < (values[0]?.length ?? 0); c += 1) {
      let run = 0
      while (run < rows && filled(values[run]?.[c])) run += 1
      seed = Math.min(seed, Math.max(1, run))
    }
    if (seed >= rows) return
    await runtime.univerAPI.executeCommand(AUTO_FILL_COMMAND, {
      sourceRange: { ...range, endRow: range.startRow + seed - 1 },
      targetRange: range,
      unitId: selection.unitId,
      subUnitId: selection.sheetId,
      applyType: APPLY_SERIES,
    })
    return
  }
  const columns = values[0]?.length ?? 0
  let seed = columns
  for (let r = 0; r < values.length; r += 1) {
    let run = 0
    while (run < columns && filled(values[r]?.[run])) run += 1
    seed = Math.min(seed, Math.max(1, run))
  }
  if (seed >= columns) return
  await runtime.univerAPI.executeCommand(AUTO_FILL_COMMAND, {
    sourceRange: { ...range, endColumn: range.startColumn + seed - 1 },
    targetRange: range,
    unitId: selection.unitId,
    subUnitId: selection.sheetId,
    applyType: APPLY_SERIES,
  })
}

/// Home → Fill → Series… OK
export async function fillSeries(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  params: FillSeriesParams,
): Promise<void> {
  const selection = currentSelection(runtime)
  const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!selection || !sheet) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const { range } = selection
  const rect: SeriesRect = {
    startRow: range.startRow,
    startColumn: range.startColumn,
    rowCount: range.endRow - range.startRow + 1,
    columnCount: range.endColumn - range.startColumn + 1,
  }
  const streamed = ctx.lazyWorkbookRef.current !== null
  if (streamed) {
    const loaded = await ensureRangeResident(
      runtime,
      ctx.lazyWorkbookRef,
      sheet,
      range,
      ctx.setMessage,
    )
    if (!loaded) {
      ctx.setMessage(t('appAutofillStreaming'))
      return
    }
  }
  const values = sheet
    .getRange(range.startRow, range.startColumn, rect.rowCount, rect.columnCount)
    .getRawValues() as unknown[][]

  if (params.type === 'autofill') {
    await autofillSeries(runtime, selection, params.axis, values)
    return
  }

  const options: SeriesOptions = {
    axis: params.axis,
    type: params.type,
    dateUnit: params.dateUnit,
    step: params.step,
    stop: params.stop,
    trend: params.trend,
    date1904: ctx.lazyWorkbookRef.current?.file.date1904 === true,
  }
  const plan = planSeries(
    rect,
    options,
    (row, column) => cellNumber(values[row - rect.startRow]?.[column - rect.startColumn]),
    { maxRows: sheet.getMaxRows(), maxColumns: sheet.getMaxColumns() },
  )
  if (plan.skipped) {
    ctx.setMessage(t(SKIP_MESSAGES[plan.skipped]))
    return
  }
  if (plan.writes.length === 0) return

  // a single cell with a Stop value runs past the selection: those rows must be loaded too
  if (streamed) {
    // (loops, not Math.min(...rows): a long series would overflow the argument list)
    const reach: IRange = { ...range }
    for (const write of plan.writes) {
      if (write.row < reach.startRow) reach.startRow = write.row
      if (write.row > reach.endRow) reach.endRow = write.row
      if (write.column < reach.startColumn) reach.startColumn = write.column
      if (write.column > reach.endColumn) reach.endColumn = write.column
    }
    const loaded = await ensureRangeResident(
      runtime,
      ctx.lazyWorkbookRef,
      sheet,
      reach,
      ctx.setMessage,
    )
    if (!loaded) {
      ctx.setMessage(t('appAutofillStreaming'))
      return
    }
  }

  // Excel gives every cell of a line the format of the line's first cell. A style object merges
  // into the old one key by key, so the old format is cleared first ({ s: null }), then the first
  // cell's goes on — otherwise a percent format further down the column would survive.
  const styleOf = new Map<string, IStyleData | null>()
  for (const start of plan.starts) {
    styleOf.set(
      `${start.row}:${start.column}`,
      sheet.getRange(start.row, start.column, 1, 1).getCellStyleData() ?? null,
    )
  }
  const lineStart = (row: number, column: number): string =>
    params.axis === 'columns' ? `${rect.startRow}:${column}` : `${row}:${rect.startColumn}`
  const clear: Record<number, Record<number, ICellData>> = {}
  const apply: Record<number, Record<number, ICellData>> = {}
  for (const write of plan.writes) {
    const style = styleOf.get(lineStart(write.row, write.column)) ?? null
    ;(clear[write.row] ??= {})[write.column] = { s: null } as unknown as ICellData
    ;(apply[write.row] ??= {})[write.column] = {
      v: write.value,
      t: CellValueType.NUMBER,
      // a formula or rich text the old cell carried must not survive under the new constant
      f: null,
      si: null,
      p: null,
      ...(style ? { s: style } : {}),
    }
  }
  const unitId = selection.unitId
  const subUnitId = selection.sheetId
  await runAsOneUndoStepAsync(runtime, async () => {
    const run = (value: unknown) =>
      runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
        unitId,
        subUnitId,
        value,
      })
    await run(clear)
    await run(apply)
  })
}
