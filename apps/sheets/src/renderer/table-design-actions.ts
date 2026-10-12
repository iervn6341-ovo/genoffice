/**
 * Table Design tab: Table Name, Table Style Options (banded rows / columns, first / last column,
 * Header Row, Total Row, Filter Button), Resize Table, the style gallery and Convert to Range on a
 * table this session created. Every change is the same move Format as Table makes over an existing
 * table — the Univer table is removed and a new one added under the same journal entry, together
 * with the cell changes it needs — so it is one ⌘Z step and the journal, the drawn table and the
 * saved tableStyleInfo cannot disagree.
 */
import {
  IUniverInstanceService,
  UniverInstanceType,
  type ICellData,
  type IRange,
  type IStyleData,
  type Workbook,
} from '@univerjs/core'
import { RangeThemeStyle } from '@univerjs/sheets'
import type { WorkbookTableAdd } from '../shared/desktop-api'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import { isSheetRemoved, journalSize, removeTableAdd } from './edit-journal'
import { t, type StringKey } from './i18n/locale'
import {
  checkResize,
  dataRowsOf,
  planTableChange,
  sessionTableAt,
  tableNameProblem,
  tableOptionsOf,
  univerTableRange,
  type CellWrite,
  type TableArea,
  type TableChangeError,
  type TableDesignOptions,
} from './table-design'
import {
  bakeTableStyle,
  DEFAULT_THEME_COLORS,
  plainTheme,
  tableThemeFor,
  totalRowTheme,
  withoutOwnFormatting,
  type TableThemeJson,
  type TableThemeSheet,
} from './table-styles'
import { forgetTable, tableIdOf } from './table-sync'
import { runAsOneUndoStepAsync } from './undo-group'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import { applyAiTableAdd } from './workbook-ops'

export interface TableDesignContext {
  readonly univerRef: { readonly current: UniverRuntime | null }
  readonly lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  readonly setMessage: (message: string) => void
  readonly setPendingEdits: (count: number) => void
}

/// The formatting a cell was given itself — what its own style holds, not the workbook's defaults
/// (`getCellStyleData` resolves those in, so every cell would look formatted).
function ownStyleOf(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  row: number,
  column: number,
): IStyleData | null {
  const own = worksheet.getRange(row, column, 1, 1).getCellData()?.s
  if (!own) return null
  if (typeof own !== 'string') return own
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  const workbook = unitId
    ? runtime.univer
        .__getInjector()
        .get(IUniverInstanceService)
        .getUnit<Workbook>(unitId, UniverInstanceType.UNIVER_SHEET)
    : null
  return workbook?.getStyles().get(own) ?? null
}

/// Excel's default table style, which a table saved without one is written (and drawn) in
const DEFAULT_TABLE_STYLE = 'TableStyleMedium2'

/// past this many cells Convert to Range bakes without looking at each cell's own formatting
const BAKE_KEEP_OWN_MAX_CELLS = 20_000

type TableWorksheet = UniverWorksheet & {
  removeTable(id: string): Promise<boolean>
} & TableThemeSheet

/** the session table under the active cell, with what is needed to change it */
export function activeSessionTable(
  ctx: TableDesignContext,
  /** the table by name on the active sheet, instead of the one under the active cell */
  byName?: string,
): {
  runtime: UniverRuntime
  state: LazyWorkbookState
  worksheet: TableWorksheet
  entry: WorkbookTableAdd
  tableId: string | undefined
} | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!runtime || !state || !worksheet || !active) return null
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return null
  const entry =
    byName === undefined
      ? sessionTableAt(state.editJournal.tableAdds, sheetId, active.getRow(), active.getColumn())
      : (state.editJournal.tableAdds.find(
          (table) => table.sheetId === sheetId && table.name === byName,
        ) ?? null)
  if (!entry) return null
  return {
    runtime,
    state,
    worksheet: worksheet as unknown as TableWorksheet,
    entry,
    tableId: tableIdOf(sheetId, entry.name),
  }
}

const rangeOf = (area: WorkbookTableAdd['area']): string =>
  `${columnLabel(area.startColumn)}${area.startRow + 1}:${columnLabel(area.endColumn)}${area.endRow + 1}`

async function finish(
  ctx: TableDesignContext,
  state: LazyWorkbookState,
  work: Promise<void>,
  doneKey: StringKey,
): Promise<void> {
  try {
    await work
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t(doneKey))
  } catch (error: unknown) {
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(error instanceof Error ? error.message : t('appTableCreateFailed'))
  }
}

const CHANGE_MESSAGES: Record<TableChangeError, StringKey> = {
  noRoomAbove: 'appTableNoRoomAbove',
  noRoomBelow: 'appTableNoRoomBelow',
  needsData: 'appTableNeedsData',
  resizeStart: 'appTableResizeStart',
  resizeTotals: 'appTableResizeTotals',
  resizeOutside: 'appTableResizeBadRange',
}

// ── the looks laid over a table on their own ─────────────────────────────────────────────────────
// Univer's table theme covers header + data rows with a header row at the top. A table without a
// header row (its theme is laid over the data rows) and a totals row (a row of its own) are drawn
// with range themes — display only, so nothing of them is journaled as a cell format.

const bodyThemeName = (tableId: string): string => `genoffice-body-${tableId}`
const totalThemeName = (tableId: string): string => `genoffice-total-${tableId}`

const toRange = (area: TableArea): IRange => ({
  startRow: area.startRow,
  endRow: area.endRow,
  startColumn: area.startColumn,
  endColumn: area.endColumn,
})

/** the ranges the laid themes of a table cover, by theme name */
function laidThemeRanges(
  tableId: string,
  area: TableArea,
  options: Pick<TableDesignOptions, 'headerRow' | 'totalRow'>,
): { name: string; range: IRange }[] {
  const data = dataRowsOf(area, options)
  const laid: { name: string; range: IRange }[] = []
  if (!options.headerRow) {
    laid.push({
      name: bodyThemeName(tableId),
      range: toRange({ ...area, startRow: data.first, endRow: data.last }),
    })
  }
  if (options.totalRow) {
    laid.push({
      name: totalThemeName(tableId),
      range: toRange({ ...area, startRow: area.endRow, endRow: area.endRow }),
    })
  }
  return laid
}

export async function layTheme(
  runtime: UniverRuntime,
  subUnitId: string,
  theme: TableThemeJson,
  range: IRange,
): Promise<void> {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  await runtime.univerAPI.executeCommand('sheet.command.register-worksheet-range-theme-style', {
    unitId,
    subUnitId,
    rangeThemeStyle: new RangeThemeStyle(theme.name, theme as never),
  })
  await runtime.univerAPI.executeCommand('sheet.command.set-worksheet-range-theme-style', {
    unitId,
    subUnitId,
    range,
    themeName: theme.name,
  })
}

export async function liftThemes(
  runtime: UniverRuntime,
  subUnitId: string,
  laid: readonly { name: string; range: IRange }[],
): Promise<void> {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  for (const { name, range } of laid) {
    await runtime.univerAPI.executeCommand('sheet.command.remove-worksheet-range-theme-style', {
      unitId,
      subUnitId,
      range,
      themeName: name,
    })
    await runtime.univerAPI.executeCommand('sheet.command.unregister-worksheet-range-theme-style', {
      unitId,
      subUnitId,
      themeName: name,
    })
  }
}

/** the cell changes of a plan as a `set-range-values` matrix (a text, a formula, or an emptied cell) */
function writesMatrix(writes: readonly CellWrite[]): Record<number, Record<number, ICellData>> {
  const matrix: Record<number, Record<number, ICellData>> = {}
  for (const write of writes) {
    const cell: ICellData =
      write.value === null
        ? { v: null, f: null, si: null, p: null }
        : write.value.startsWith('=')
          ? { v: null, f: write.value, si: null, p: null }
          : { v: write.value, t: 1, f: null, si: null, p: null }
    ;(matrix[write.row] ??= {})[write.column] = cell
  }
  return matrix
}

/**
 * Re-creates the active session table with a new name, style, options, header / totals rows or
 * area (one undo step). The name and the options carry over unless the patch changes them.
 */
export async function redesignTable(
  ctx: TableDesignContext,
  patch: {
    name?: string
    style?: string
    options?: Partial<TableDesignOptions>
    /** Resize Table: the new range of the table */
    area?: TableArea
    /** the table to change by its current name (a rename typed, then another table clicked) */
    target?: string
  },
  doneKey: StringKey = 'appTableRestyled',
): Promise<void> {
  const found = activeSessionTable(ctx, patch.target)
  if (!found || !found.tableId) {
    ctx.setMessage(t('appTableDesignNeedsTable'))
    return
  }
  const { runtime, state, worksheet, entry, tableId: oldTableId } = found
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const name = patch.name ?? entry.name
  // a table saved without a style is Excel's default one, and is drawn as such
  const style = patch.style ?? entry.style ?? DEFAULT_TABLE_STYLE
  const oldOptions = tableOptionsOf(entry)
  const options: TableDesignOptions = { ...oldOptions, ...patch.options }
  if (patch.name !== undefined) {
    const problem = tableNameProblem(
      name,
      [
        ...state.editJournal.tableAdds.map((table) => table.name),
        ...state.file.sheets.flatMap((sheet) => sheet.tables.map((table) => table.name ?? '')),
      ].filter((other) => other !== ''),
      entry.name,
      [
        ...(workbook?.getSheets().map((sheet) => sheet.getSheetName()) ?? []),
        ...(workbook?.getDefinedNames().map((defined) => defined.getName()) ?? []),
      ],
    )
    if (problem) {
      const key: StringKey =
        problem === 'taken'
          ? 'appTableNameUsed'
          : problem === 'reserved'
            ? 'appTableNameReserved'
            : 'appTableNameInvalid'
      ctx.setMessage(t(key, { name }))
      return
    }
  }

  // ── Resize, Header Row and Total Row change cells and the area; the rest only redraws ──
  let area: TableArea = entry.area
  let columnNames: readonly string[] = entry.columnNames
  let writes: readonly CellWrite[] = []
  let totals: WorkbookTableAdd['totals'] | undefined = entry.totals
  if (patch.area) {
    const problem = checkResize(entry.area, patch.area, oldOptions)
    if (problem || patch.area.endRow >= worksheet.getMaxRows()) {
      ctx.setMessage(t(CHANGE_MESSAGES[problem ?? 'resizeOutside']))
      return
    }
    area = patch.area
    const width = area.endColumn - area.startColumn + 1
    columnNames = Array.from({ length: width }, (_, i) => entry.columnNames[i] ?? `Column${i + 1}`)
  } else if (
    options.headerRow !== oldOptions.headerRow ||
    options.totalRow !== oldOptions.totalRow
  ) {
    const change = planTableChange({
      area: entry.area,
      columnNames: entry.columnNames,
      from: oldOptions,
      to: options,
      tableName: name,
      totalLabel: t('appTableTotalLabel'),
      maxRows: worksheet.getMaxRows(),
      cellAt: (row, column) => worksheet.getRange(row, column, 1, 1).getCellData()?.v ?? null,
    })
    if ('error' in change) {
      ctx.setMessage(t(CHANGE_MESSAGES[change.error]))
      return
    }
    area = change.area
    columnNames = change.columnNames
    writes = change.writes
    totals = options.totalRow ? (change.totals ?? entry.totals) : undefined
  }
  // a totals row that a rename touches: its formulas follow the table's new name
  if (options.totalRow && patch.name !== undefined && patch.name !== entry.name && !patch.area) {
    const lastColumn = entry.area.endColumn
    const cell = worksheet.getRange(entry.area.endRow, lastColumn, 1, 1).getCellData()
    const formula = cell?.f
    if (typeof formula === 'string') {
      writes = [
        ...writes,
        {
          row: entry.area.endRow,
          column: lastColumn,
          value: formula.replaceAll(`${entry.name}[`, `${name}[`),
        },
      ]
    }
  }

  const colors = state.file.themeColors ?? DEFAULT_THEME_COLORS
  const oldLaid = laidThemeRanges(oldTableId, entry.area, oldOptions)
  await finish(
    ctx,
    state,
    runAsOneUndoStepAsync(runtime, async () => {
      // table-sync drops the old journal entry when Univer's delete-table mutation runs
      await worksheet.removeTable(oldTableId!)
      await liftThemes(runtime, worksheet.getSheetId(), oldLaid)
      if (writes.length > 0) {
        await runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
          unitId: workbook?.getId(),
          subUnitId: worksheet.getSheetId(),
          value: writesMatrix(writes),
        })
      }
      const { tableId, ready } = applyAiTableAdd(
        runtime,
        state,
        {
          op: 'add_table',
          sheetId: entry.sheetId,
          range: rangeOf(area),
          name,
          ...(style === undefined ? {} : { style }),
          bandedRows: options.bandedRows,
        },
        options,
        { columnNames, ...(totals ? { totals } : {}) },
      )
      if (!(await ready)) {
        // Univer refused the table (a name it does not take, a range over another table): no ghost
        // in the journal, and the table that was there comes back as it was
        forgetTable(tableId)
        removeTableAdd(state.editJournal, entry.sheetId, name)
        const restored = applyAiTableAdd(
          runtime,
          state,
          {
            op: 'add_table',
            sheetId: entry.sheetId,
            range: rangeOf(entry.area),
            name: entry.name,
            ...(entry.style === undefined ? {} : { style: entry.style }),
            bandedRows: entry.bandedRows,
          },
          oldOptions,
          { columnNames: entry.columnNames, ...(entry.totals ? { totals: entry.totals } : {}) },
        )
        await restored.ready
        const oldTheme = tableThemeFor(
          entry.style ?? DEFAULT_TABLE_STYLE,
          `genoffice-${restored.tableId}`,
          colors,
          oldOptions,
        )
        if (oldTheme) {
          await worksheet.addTableTheme(
            restored.tableId,
            oldOptions.headerRow ? oldTheme : plainTheme(`genoffice-${restored.tableId}`),
          )
          await layDefinedThemes(
            runtime,
            worksheet,
            restored.tableId,
            entry.style ?? DEFAULT_TABLE_STYLE,
            entry.area,
            oldOptions,
            colors,
          )
        }
        throw new Error(t('appTableResizeRefused', { name: entry.name }))
      }
      const theme = tableThemeFor(style, `genoffice-${tableId}`, colors, options)
      if (theme) {
        // a table without a header row keeps Univer's table for its filter / formula bookkeeping
        // only; the look is laid over the data rows
        await worksheet.addTableTheme(
          tableId,
          options.headerRow ? theme : plainTheme(`genoffice-${tableId}`),
        )
        await layDefinedThemes(runtime, worksheet, tableId, style, area, options, colors)
      }
    }),
    patch.area
      ? 'appTableResized'
      : options.headerRow !== oldOptions.headerRow ||
          options.totalRow !== oldOptions.totalRow ||
          options.filterButton !== oldOptions.filterButton
        ? 'appTableOptionChanged'
        : doneKey,
  )
}

/** lays the body theme (a table without a header row) and the totals row theme over their rows */
async function layDefinedThemes(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  tableId: string,
  style: string,
  area: TableArea,
  options: TableDesignOptions,
  colors: readonly string[],
): Promise<void> {
  for (const { name, range } of laidThemeRanges(tableId, area, options)) {
    const theme = name.startsWith('genoffice-body-')
      ? tableThemeFor(style, name, colors, options)
      : totalRowTheme(style, name, colors, options)
    if (theme) await layTheme(runtime, worksheet.getSheetId(), theme, range)
  }
}

/**
 * Convert to Range: the table object goes, its look stays — baked into the cells' own formats,
 * which is what Excel does — in one undo step.
 */
export async function convertTableToRange(ctx: TableDesignContext): Promise<void> {
  const found = activeSessionTable(ctx)
  if (!found || !found.tableId) {
    ctx.setMessage(t('appTableDesignNeedsTable'))
    return
  }
  const { runtime, state, worksheet, entry, tableId } = found
  const options = tableOptionsOf(entry)
  // a table saved without a style is Excel's default one; its look is baked too
  const baked = bakeTableStyle(
    entry.style ?? DEFAULT_TABLE_STYLE,
    entry.area,
    state.file.themeColors ?? DEFAULT_THEME_COLORS,
    options,
  )
  // the cells' own formatting stays above the style (a yellow cell stays yellow, as in Excel)
  const cells =
    (entry.area.endRow - entry.area.startRow + 1) *
    (entry.area.endColumn - entry.area.startColumn + 1)
  if (baked && cells <= BAKE_KEEP_OWN_MAX_CELLS) {
    for (const [rowKey, line] of Object.entries(baked)) {
      for (const [columnKey, cell] of Object.entries(line)) {
        const own = ownStyleOf(runtime, worksheet, Number(rowKey), Number(columnKey))
        const kept = withoutOwnFormatting(cell.s, own)
        if (kept) cell.s = kept
        else delete line[Number(columnKey)]
      }
    }
  }
  await finish(
    ctx,
    state,
    runAsOneUndoStepAsync(runtime, async () => {
      // table-sync drops the journal entry when Univer's delete-table mutation runs
      await worksheet.removeTable(tableId!)
      // the display-only looks laid over the table give way to the baked ones
      await liftThemes(
        runtime,
        worksheet.getSheetId(),
        laidThemeRanges(tableId!, entry.area, options),
      )
      if (baked) {
        const workbook = runtime.univerAPI.getActiveWorkbook()
        await runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
          unitId: workbook?.getId(),
          subUnitId: worksheet.getSheetId(),
          value: baked,
        })
      }
    }),
    'appTableConverted',
  )
}
