/**
 * Table Design on tables the file already has: the style gallery and the four Table Style Options
 * (banded rows / columns, first / last column). The save patches the table part's
 * <tableStyleInfo> (`fileTableEdits` in the journal); everything else of the part — columns,
 * filter, totals row, the ref — stays as the file has it, which is why the structural controls
 * (rename, resize, Header Row, Total Row, Filter Button, Convert to Range) are for session tables.
 *
 * The file's look is baked into the cells' formats at load (applyTableBanding), so a restyle takes
 * that baked look off the cells that are loaded and lets display-only range themes draw the new
 * one; the table's descriptor is marked so later streaming does not bake the old look again. Undo
 * and redo put the look back through a journal undo entry that holds the *effective* look, so they
 * also work across a Save, which reopens the workbook with the new style as the file's own.
 */
import {
  IUniverInstanceService,
  UniverInstanceType,
  type ICellData,
  type IRange,
  type IStyleData,
  type Workbook,
} from '@univerjs/core'

import type { WorkbookFile, WorkbookTableEdit, WorkbookTableInfo } from '../shared/desktop-api'
import { isSheetRemoved, journalSize, recordFileTableEdit } from './edit-journal'
import { t, type StringKey } from './i18n/locale'
import { showToast } from './toast-bus'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'
import {
  activeSessionTable,
  layTheme,
  liftThemes,
  redesignTable,
  type TableDesignContext,
} from './table-design-actions'
import type { TableArea, TableDesignOptions } from './table-design'
import { DEFAULT_THEME_COLORS, tableThemeFor, totalRowTheme } from './table-styles'
import { applyTableBanding, hiddenRowsInfo } from './univer-sync'
import {
  journalSuppression,
  lazyFileSheetId,
  lazySheetMeta,
  restyledFileTables,
  type FileTableLook,
  type LazyWorkbookState,
  type UniverRuntime,
} from './univer-state'
import { fileRangeToScreenRange } from './view-transform'

type FileTable = WorkbookFile['sheets'][number]['tables'][number]

export interface FileTableRef {
  /** sheet + table name: what the state maps are keyed by */
  readonly key: string
  readonly sheetId: string
  readonly name: string
  /** the table's place in the sheet's table list (its Univer table is registered under it) */
  readonly index: number
  readonly table: FileTable
  /** the table's cells; they are where the file has them (a table that rows / columns edits moved is left alone) */
  readonly area: TableArea
}

export interface FileTableContext extends TableDesignContext {
  readonly requestRender?: () => void
}

/// Re-renders the ribbon: undo and redo restore the look asynchronously and reach no component.
let renderNudge: (() => void) | null = null
export function setFileTableRenderer(nudge: (() => void) | null): void {
  renderNudge = nudge
}

/// the loaded cells of a table that a restyle repaints, at most (a table of a huge file is left alone)
const RESTYLE_CELL_CAP = 500_000

const keyOf = (sheetId: string, name: string): string => `${sheetId}\u0000${name.toLowerCase()}`

const sameArea = (a: TableArea, b: TableArea): boolean =>
  a.startRow === b.startRow &&
  a.endRow === b.endRow &&
  a.startColumn === b.startColumn &&
  a.endColumn === b.endColumn

/** the named tables of a sheet that sit where the file has them (no structural edit moved them) */
function fileTablesOf(state: LazyWorkbookState, sheetId: string): FileTableRef[] {
  // a duplicated sheet streams from another one; its tables are not part of the file
  if (lazyFileSheetId(state, sheetId) !== sheetId) return []
  const meta = lazySheetMeta(state, sheetId)
  if (!meta || isSheetRemoved(state.editJournal, sheetId)) return []
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const refs: FileTableRef[] = []
  for (const [index, table] of meta.tables.entries()) {
    if (table.name === undefined || table.name === '') continue
    const moved = ops.length === 0 ? table.range : fileRangeToScreenRange(ops, table.range)
    if (!moved || !sameArea(moved, table.range)) continue
    refs.push({
      key: keyOf(sheetId, table.name),
      sheetId,
      name: table.name,
      index,
      table,
      area: table.range,
    })
  }
  return refs
}

/** the file's table under a cell, or null */
export function fileTableAt(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): FileTableRef | null {
  return (
    fileTablesOf(state, sheetId).find(
      ({ area }) =>
        row >= area.startRow &&
        row <= area.endRow &&
        column >= area.startColumn &&
        column <= area.endColumn,
    ) ?? null
  )
}

const lookOfInfo = (info: WorkbookTableInfo): FileTableLook => ({
  style: info.style ?? null,
  bandedRows: info.showRowStripes,
  bandedColumns: info.showColumnStripes,
  firstColumn: info.showFirstColumn,
  lastColumn: info.showLastColumn,
})

/// the descriptor knows the style and the stripes; first / last column wait for the table part
const lookOfDescriptor = (table: FileTable): FileTableLook => ({
  style: table.styleName ?? null,
  bandedRows: table.showRowStripes,
  bandedColumns: table.showColumnStripes,
  firstColumn: false,
  lastColumn: false,
})

const sameLook = (a: FileTableLook, b: FileTableLook): boolean =>
  a.style === b.style &&
  a.bandedRows === b.bandedRows &&
  a.bandedColumns === b.bandedColumns &&
  a.firstColumn === b.firstColumn &&
  a.lastColumn === b.lastColumn

/** what the table part of the file says (the descriptor's guess until the part has been read) */
export function fileLookOf(state: LazyWorkbookState, ref: FileTableRef): FileTableLook {
  const info = state.fileTableInfos.get(ref.name.toLowerCase())
  return info ? lookOfInfo(info) : lookOfDescriptor(ref.table)
}

/** what the table looks like now: the look Table Design gave it this session, else the file's */
export function effectiveLookOf(state: LazyWorkbookState, ref: FileTableRef): FileTableLook {
  return state.fileTableShown.get(ref.key)?.look ?? fileLookOf(state, ref)
}

/** the options Table Design shows: the structural ones are what the file's table has */
export function designOptionsOf(ref: FileTableRef, look: FileTableLook): TableDesignOptions {
  const headerRow = ref.table.headerRowCount > 0
  return {
    bandedRows: look.bandedRows,
    bandedColumns: look.bandedColumns,
    firstColumn: look.firstColumn,
    lastColumn: look.lastColumn,
    headerRow,
    totalRow: (ref.table.totalsRowCount ?? 0) > 0,
    filterButton: headerRow,
  }
}

/** the table part's style info, read once per table (the Table Design tab asks for it when it opens) */
export async function ensureFileTableInfo(
  state: LazyWorkbookState,
  name: string,
): Promise<WorkbookTableInfo | null> {
  const key = name.toLowerCase()
  const known = state.fileTableInfos.get(key)
  if (known) return known
  try {
    const info = await window.desktopApi.readTableInfo({ sessionId: state.file.sessionId, name })
    state.fileTableInfos.set(key, info)
    return info
  } catch {
    return null
  }
}

/** asks for the table's info in the background; `onLoaded` re-renders what shows it */
export function prefetchFileTableInfo(
  state: LazyWorkbookState,
  name: string,
  onLoaded: () => void,
): void {
  const key = name.toLowerCase()
  if (state.fileTableInfos.has(key) || state.fileTableInfoRequested.has(key)) return
  state.fileTableInfoRequested.add(key)
  // a read that failed is not asked again on every render: the next action asks once more
  void ensureFileTableInfo(state, name).then((info) => {
    if (info) onLoaded()
  })
}

// ── drawing the look ─────────────────────────────────────────────────────────────────────────────

/// whether the grid can draw a style: the built-in ones and the custom ones made in this workbook; a
/// style the file defines itself (a corporate template) has no definition here
function canDraw(style: string | null, colors: readonly string[]): boolean {
  return style === null || tableThemeFor(style, 'genoffice-probe', colors) !== null
}

const rangeOf = (area: TableArea): IRange => ({
  startRow: area.startRow,
  endRow: area.endRow,
  startColumn: area.startColumn,
  endColumn: area.endColumn,
})

/// a cell as Univer holds it, through the slice of the worksheet this module reads
interface CellMatrixLike {
  getValue(row: number, column: number): ICellData | null | undefined
}

/** reads a cell's own format (the workbook's styles resolve a style id) */
function styleReader(
  runtime: UniverRuntime,
  unitId: string,
): (cell: ICellData | null | undefined) => IStyleData | null {
  const workbook = runtime.univer
    .__getInjector()
    .get(IUniverInstanceService)
    .getUnit<Workbook>(unitId, UniverInstanceType.UNIVER_SHEET)
  const styles = workbook?.getStyles()
  return (cell) => {
    const own = cell?.s
    if (!own) return null
    return typeof own === 'string' ? (styles?.get(own) ?? null) : own
  }
}

const ATTRS = ['bg', 'cl', 'bl'] as const
const EDGES = ['t', 'b', 'l', 'r'] as const

const sameValue = (a: unknown, b: unknown): boolean =>
  JSON.stringify(a ?? null).toLowerCase() === JSON.stringify(b ?? null).toLowerCase()

/** the table's cells that are loaded in the grid (the file's cells the streaming installed) */
function loadedBox(state: LazyWorkbookState, ref: FileTableRef): IRange | null {
  const loaded = state.loadedRanges.get(ref.sheetId)
  if (!loaded) return null
  const box: IRange = {
    startRow: Math.max(loaded.startRow, ref.area.startRow),
    endRow: Math.min(loaded.endRow, ref.area.endRow),
    startColumn: Math.max(loaded.startColumn, ref.area.startColumn),
    endColumn: Math.min(loaded.endColumn, ref.area.endColumn),
  }
  if (box.startRow > box.endRow || box.startColumn > box.endColumn) return null
  const cells = (box.endRow - box.startRow + 1) * (box.endColumn - box.startColumn + 1)
  if (cells > RESTYLE_CELL_CAP) throw new Error(t('appFileTableTooBig'))
  return box
}

/**
 * Takes the baked look of the table off the loaded cells (what applyTableBanding painted: only
 * the attributes that still equal it, so a fill or font the cell has of its own stays), or — with
 * `put` — paints it back on the cells that have no formatting of their own for it.
 */
async function reshapeBakedLook(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  ref: FileTableRef,
  mode: 'strip' | 'put',
): Promise<void> {
  const box = loadedBox(state, ref)
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(ref.sheetId)
  if (!box || !workbook || !worksheet) return
  const unitId = workbook.getId()
  const read = styleReader(runtime, unitId)
  const cells = (worksheet as unknown as { getSheet(): { getCellMatrix(): CellMatrixLike } })
    .getSheet()
    .getCellMatrix()
  const rows = box.endRow - box.startRow + 1
  const columns = box.endColumn - box.startColumn + 1
  const current = (row: number, column: number): IStyleData | null =>
    read(cells.getValue(box.startRow + row, box.startColumn + column))
  // what the cells would be painted with: the loaded formats under applyTableBanding
  const matrix: ICellData[][] = Array.from({ length: rows }, (_row, rowOffset) =>
    Array.from({ length: columns }, (_cell, columnOffset) => {
      const style = mode === 'put' ? current(rowOffset, columnOffset) : null
      return (style ? { s: structuredClone(style) } : {}) as ICellData
    }),
  )
  applyTableBanding(matrix, box, [ref.table], hiddenRowsInfo(state, ref.sheetId))
  const changes: Record<number, Record<number, ICellData>> = {}
  for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
      const painted = matrix[rowOffset]?.[columnOffset]?.s as IStyleData | undefined
      if (!painted) continue
      const own = current(rowOffset, columnOffset)
      const diff: Record<string, unknown> = {}
      for (const attr of ATTRS) {
        if (painted[attr] === undefined) continue
        if (
          mode === 'strip'
            ? sameValue(own?.[attr], painted[attr])
            : !sameValue(own?.[attr], painted[attr])
        ) {
          diff[attr] = mode === 'strip' ? null : painted[attr]
        }
      }
      const edges: Record<string, unknown> = {}
      for (const edge of EDGES) {
        const paintedEdge = painted.bd?.[edge]
        if (!paintedEdge) continue
        const ownEdge = own?.bd?.[edge]
        if (mode === 'strip' ? sameValue(ownEdge, paintedEdge) : !sameValue(ownEdge, paintedEdge)) {
          edges[edge] = mode === 'strip' ? null : paintedEdge
        }
      }
      if (Object.keys(edges).length > 0) diff.bd = edges
      if (Object.keys(diff).length === 0) continue
      ;(changes[box.startRow + rowOffset] ??= {})[box.startColumn + columnOffset] = {
        s: diff,
      } as unknown as ICellData
    }
  }
  if (Object.keys(changes).length === 0) return
  await runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
    unitId,
    subUnitId: ref.sheetId,
    value: changes,
  })
}

/** the range themes that draw `look` over a table: the header + data rows, and the totals row on its own */
function themesFor(
  ref: FileTableRef,
  look: FileTableLook,
  colors: readonly string[],
  stamp: string,
): { range: IRange; theme: NonNullable<ReturnType<typeof tableThemeFor>> }[] {
  if (look.style === null) return []
  const options = designOptionsOf(ref, look)
  const totals = ref.table.totalsRowCount ?? 0
  const bodyEnd = ref.area.endRow - totals
  const made: { range: IRange; theme: NonNullable<ReturnType<typeof tableThemeFor>> }[] = []
  const base = `genoffice-file-${stamp}`
  if (bodyEnd >= ref.area.startRow) {
    const theme = tableThemeFor(look.style, `${base}-body`, colors, options)
    if (theme) made.push({ range: rangeOf({ ...ref.area, endRow: bodyEnd }), theme })
  }
  if (totals > 0) {
    const theme = totalRowTheme(look.style, `${base}-total`, colors, options)
    if (theme) made.push({ range: rangeOf({ ...ref.area, startRow: bodyEnd + 1 }), theme })
  }
  return made
}

let themeStamp = 0

/// The plain theme the load puts over a file's table (its own look is baked into the cells). Univer
/// draws only one theme per cell — the first rule it finds — so it has to be out of the way while
/// the range themes of Table Design draw the table, and back when the file's own look is.
async function setPlainRule(
  runtime: UniverRuntime,
  ref: FileTableRef,
  present: boolean,
): Promise<void> {
  // a table without a header row was never registered with Univer
  if (ref.table.headerRowCount === 0) return
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  await runtime.univerAPI.executeCommand(
    present
      ? 'sheet.command.set-worksheet-range-theme-style'
      : 'sheet.command.remove-worksheet-range-theme-style',
    {
      unitId,
      subUnitId: ref.sheetId,
      range: rangeOf(ref.area),
      themeName: `plain-file-table-${ref.sheetId}-${ref.index}`,
    },
  )
}

/// Paints run one after the other: each holds journalSuppression for its awaits, so a second one
/// (an undo pressed while a restyle is still drawing) must not start, and lower it, in the middle.
let paintQueue: Promise<void> = Promise.resolve()

/**
 * Shows `look` on the table: null is the file's own look (baked into the cells). Safe to run again
 * for the look that is already showing — it only brings the grid to the state the look calls for.
 */
function showLook(
  ctx: JournalUndoContext,
  state: LazyWorkbookState,
  ref: FileTableRef,
  look: FileTableLook | null,
): Promise<void> {
  const run = paintQueue.then(() => showLookNow(ctx, state, ref, look))
  paintQueue = run.catch(() => undefined)
  return run
}

async function showLookNow(
  ctx: JournalUndoContext,
  state: LazyWorkbookState,
  ref: FileTableRef,
  wanted: FileTableLook | null,
): Promise<void> {
  const runtime = ctx.univerRef.current
  if (!runtime || ctx.lazyWorkbookRef.current !== state) return
  let look = wanted
  // too big to repaint: say so before anything is lifted or stripped
  loadedBox(state, ref)
  const colors = state.file.themeColors ?? DEFAULT_THEME_COLORS
  // a style that is not drawn here leaves the table as the file's own look draws it
  if (look !== null && !canDraw(look.style, colors)) look = null
  const shown = state.fileTableShown.get(ref.key)
  journalSuppression.active = true
  try {
    if (shown && shown.laid.length > 0) await liftThemes(runtime, ref.sheetId, shown.laid)
    if (look === null) {
      state.fileTableShown.delete(ref.key)
      if (restyledFileTables.has(ref.table)) {
        restyledFileTables.delete(ref.table)
        await reshapeBakedLook(runtime, state, ref, 'put')
        await setPlainRule(runtime, ref, true)
      }
      return
    }
    if (!restyledFileTables.has(ref.table)) {
      await reshapeBakedLook(runtime, state, ref, 'strip')
      await setPlainRule(runtime, ref, false)
      restyledFileTables.add(ref.table)
    }
    themeStamp += 1
    const laid: { name: string; range: IRange }[] = []
    for (const { range, theme } of themesFor(ref, look, colors, `${themeStamp}`)) {
      await layTheme(runtime, ref.sheetId, theme, range)
      laid.push({ name: theme.name, range })
    }
    state.fileTableShown.set(ref.key, { look, laid })
  } finally {
    journalSuppression.active = false
  }
}

// ── the journal, and undo ────────────────────────────────────────────────────────────────────────

/// what the table part has to change to say `look`: only the fields that differ from the file's own, so
/// a field nobody touched keeps whatever the file (or a later Excel session) has
function editFields(
  look: FileTableLook,
  inFile: FileTableLook,
): Omit<WorkbookTableEdit, 'sheetId' | 'name'> {
  return {
    ...(look.style === null || look.style === inFile.style ? {} : { style: look.style }),
    ...(look.bandedRows === inFile.bandedRows ? {} : { bandedRows: look.bandedRows }),
    ...(look.bandedColumns === inFile.bandedColumns ? {} : { bandedColumns: look.bandedColumns }),
    ...(look.firstColumn === inFile.firstColumn ? {} : { firstColumn: look.firstColumn }),
    ...(look.lastColumn === inFile.lastColumn ? {} : { lastColumn: look.lastColumn }),
  }
}

/** the table of `key` as the file has it now; null when it is gone (a deleted sheet, a converted table) */
function refOfKey(state: LazyWorkbookState, key: string): FileTableRef | null {
  const sheetId = key.slice(0, key.indexOf('\u0000'))
  return fileTablesOf(state, sheetId).find((ref) => ref.key === key) ?? null
}

/**
 * Makes `look` the table's look: what the journal saves, and what the grid shows. A look that is
 * what the table part already says leaves the journal and the grid as the file has them.
 */
async function applyLook(
  ctx: JournalUndoContext,
  state: LazyWorkbookState,
  ref: FileTableRef,
  wanted: FileTableLook,
): Promise<void> {
  // without the table part's own values the difference to write would be a guess
  if (!(await ensureFileTableInfo(state, ref.name))) throw new Error(t('appFileTableInfoFailed'))
  const inFile = fileLookOf(state, ref)
  // a table saved without a style cannot be given "none" back: the style stays the file's
  const look: FileTableLook = { ...wanted, style: wanted.style ?? inFile.style }
  const editKey = ref.name.toLowerCase()
  const same = sameLook(look, inFile)
  // the grid first: a repaint that cannot be done leaves the journal as it was
  await showLook(ctx, state, ref, same ? null : look)
  state.editJournal.fileTableEdits.delete(editKey)
  if (same) return
  recordFileTableEdit(state.editJournal, {
    sheetId: ref.sheetId,
    name: ref.name,
    ...editFields(look, inFile),
  })
}

interface LookSnapshot {
  /** the look the table had (what showed), null when the table is not known */
  look: FileTableLook | null
}

const fileTableKind: JournalUndoKind<LookSnapshot> = {
  id: 'genoffice.mutation.set-file-table-look',
  capture: (state, key) => {
    const ref = refOfKey(state, key)
    return { look: ref ? structuredClone(effectiveLookOf(state, ref)) : null }
  },
  restore: (state, key, snapshot, ctx) => {
    const ref = refOfKey(state, key)
    if (!ref || !snapshot.look) return
    void applyLook(ctx, state, ref, snapshot.look)
      .then(() => {
        ctx.setPendingEdits(journalSize(state.editJournal))
        ctx.requestRender?.()
        renderNudge?.()
      })
      .catch((error: unknown) => {
        showToast(error instanceof Error ? error.message : t('appTableCreateFailed'), 'error')
      })
  },
}

// ── the actions ──────────────────────────────────────────────────────────────────────────────────

/** the table of the file under the active cell, with what is needed to change it */
export function activeFileTable(
  ctx: FileTableContext,
): { runtime: UniverRuntime; state: LazyWorkbookState; ref: FileTableRef } | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const sheetId = workbook?.getActiveSheet()?.getSheetId()
  const active = workbook?.getActiveRange()
  if (!runtime || !state || !sheetId || !active) return null
  const ref = fileTableAt(state, sheetId, active.getRow(), active.getColumn())
  return ref ? { runtime, state, ref } : null
}

export interface FileTablePatch {
  style?: string
  options?: Partial<
    Pick<TableDesignOptions, 'bandedRows' | 'bandedColumns' | 'firstColumn' | 'lastColumn'>
  >
}

/**
 * Table Design on the file's table under the active cell: a style and / or style options, one undo
 * step. Returns false when there is no such table (the caller then tries a session table).
 */
export async function restyleFileTable(
  ctx: FileTableContext,
  patch: FileTablePatch,
  doneKey: StringKey = 'appTableRestyled',
): Promise<boolean> {
  const found = activeFileTable(ctx)
  if (!found) return false
  const { state, ref } = found
  try {
    if (!(await ensureFileTableInfo(state, ref.name))) {
      ctx.setMessage(t('appFileTableInfoFailed'))
      return true
    }
    const before = effectiveLookOf(state, ref)
    const next: FileTableLook = {
      ...before,
      ...(patch.style === undefined ? {} : { style: patch.style }),
      ...patch.options,
    }
    if (sameLook(before, next)) return true
    // a style the file defines itself cannot be drawn here: an option of it would only blank the table
    if (!canDraw(next.style, state.file.themeColors ?? DEFAULT_THEME_COLORS)) {
      ctx.setMessage(t('appFileTableStyleUnknown'))
      return true
    }
    await withJournalUndo(ctx, fileTableKind, ref.key, () => applyLook(ctx, state, ref, next))
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t(doneKey))
    ctx.requestRender?.()
    renderNudge?.()
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appTableCreateFailed'))
  }
  return true
}

/**
 * Table Design / Format as Table on the table under the active cell, whichever kind it is: a
 * session table is re-created (redesignTable), a table of the file gets its style and options.
 */
export async function restyleActiveTable(
  ctx: FileTableContext,
  patch: { style?: string; options?: Partial<TableDesignOptions> },
  doneKey: StringKey = 'appTableRestyled',
): Promise<void> {
  if (!activeSessionTable(ctx) && activeFileTable(ctx)) {
    const { headerRow, totalRow, filterButton, ...banding } = patch.options ?? {}
    // the structural options belong to the table part, which a file's table keeps as it is
    if (headerRow !== undefined || totalRow !== undefined || filterButton !== undefined) {
      ctx.setMessage(t('appFileTableLocked'))
      return
    }
    await restyleFileTable(
      ctx,
      { ...(patch.style === undefined ? {} : { style: patch.style }), options: banding },
      doneKey,
    )
    return
  }
  await redesignTable(ctx, patch, doneKey)
}
