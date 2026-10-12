/**
 * Home → Clear: the parts of Excel's Clear menu that Univer's clear commands do not reach — notes
 * and hyperlinks. Excel (16, probed through COM) draws the lines like this:
 *   Clear All                  text, formats, notes and hyperlinks
 *   Clear Contents             text and hyperlinks; formats and notes stay
 *   Clear Formats              formats only
 *   Clear Comments and Notes   notes only
 *   Clear Hyperlinks           the link; text, note and the link's blue underline stay
 *   Remove Hyperlinks          the link and its formatting
 * Hyperlinks live in the edit journal (the file's own links stream in lazily), so an undo entry
 * restores the journal side; the formatting side rides on Univer's own undo.
 */
import type { ICellData, IRange } from '@univerjs/core'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import { hyperlinkEditAt, journalSize, recordHyperlinkEdit } from './edit-journal'
import { withJournalUndo, type JournalUndoKind } from './journal-undo'
import type { RibbonCommandContext } from './ribbon-actions'
import { deleteNotesUndoable } from './note-undo'
import { runAsOneUndoStepAsync } from './undo-group'
import { ensureRangeResident } from './lazy-resident'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// A selection bigger than this is not loaded in full before a clear looks for its links / notes
/// (Ctrl+A over a streamed file would stream every row first); the ones already loaded are cleared.
const LOAD_BEFORE_CLEAR_MAX_CELLS = 250_000

const cellCount = (range: IRange): number =>
  (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)

/// the effective link target of the cells a clear touched: null = no link
type HyperlinkSnapshot = readonly (readonly [string, string | null])[]

interface HyperlinkKey {
  readonly sheetId: string
  /// the "row:column" keys of the cells the clear touches
  readonly cells: readonly string[]
}

const effectiveTarget = (
  state: LazyWorkbookState,
  sheetId: string,
  cell: string,
): string | null => {
  const [row, column] = cell.split(':').map(Number)
  const edited = hyperlinkEditAt(state.editJournal, sheetId, row ?? 0, column ?? 0)
  return edited !== undefined ? edited : (state.hyperlinkTargets.get(sheetId)?.get(cell) ?? null)
}

/**
 * A hyperlink clear journals link edits and drops streamed targets; ⌘Z has to put the links back.
 * The snapshot is each touched cell's EFFECTIVE target, and restoring writes it as an edit again —
 * not "the journal as it was": a save in between reopens the workbook from the file that already
 * lacks the links, with an empty journal, so rolling the journal back would leave the file without
 * them and nothing pending to write them.
 */
const hyperlinkKind: JournalUndoKind<HyperlinkSnapshot> = {
  id: 'genoffice.mutation.set-hyperlinks',
  capture: (state, encoded) => {
    const key = JSON.parse(encoded) as HyperlinkKey
    return key.cells.map((cell) => [cell, effectiveTarget(state, key.sheetId, cell)] as const)
  },
  restore: (state, encoded, snapshot, ctx) => {
    const key = JSON.parse(encoded) as HyperlinkKey
    const targets = state.hyperlinkTargets.get(key.sheetId) ?? new Map<string, string>()
    state.hyperlinkTargets.set(key.sheetId, targets)
    for (const [cell, target] of snapshot) {
      const [row, column] = cell.split(':').map(Number)
      recordHyperlinkEdit(state.editJournal, key.sheetId, row ?? 0, column ?? 0, target)
      if (target === null) targets.delete(cell)
      else targets.set(cell, target)
    }
    ctx.setPendingEdits(journalSize(state.editJournal))
  },
}

/// The sheet's cells with a hyperlink inside `range` (the journal's edit wins over the file's target)
export function hyperlinkCellsIn(
  state: LazyWorkbookState,
  sheetId: string,
  range: IRange,
): { row: number; column: number }[] {
  const keys = new Set<string>([
    ...(state.hyperlinkTargets.get(sheetId)?.keys() ?? []),
    ...(state.editJournal.hyperlinks.get(sheetId)?.keys() ?? []),
  ])
  const cells: { row: number; column: number }[] = []
  for (const key of keys) {
    const [row, column] = key.split(':').map(Number)
    if (row === undefined || column === undefined) continue
    if (row < range.startRow || row > range.endRow) continue
    if (column < range.startColumn || column > range.endColumn) continue
    const edited = hyperlinkEditAt(state.editJournal, sheetId, row, column)
    const target = edited !== undefined ? edited : state.hyperlinkTargets.get(sheetId)?.get(key)
    if (target) cells.push({ row, column })
  }
  return cells
}

function activeSelection(runtime: UniverRuntime): {
  sheetId: string
  range: IRange
} | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!sheet || !active) return null
  return {
    sheetId: sheet.getSheetId(),
    range: {
      startRow: active.getRow(),
      endRow: active.getRow() + active.getHeight() - 1,
      startColumn: active.getColumn(),
      endColumn: active.getColumn() + active.getWidth() - 1,
    },
  }
}

/**
 * Drops the hyperlinks under the selection. `removeFormat` also takes off the blue underline that
 * Excel's Hyperlink style carries (Remove Hyperlinks); without it the cell keeps its look (Clear
 * Hyperlinks, Clear Contents). Returns how many links went.
 */
export async function clearHyperlinks(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  options: { removeFormat: boolean },
): Promise<number> {
  const state = ctx.lazyWorkbookRef.current
  const selection = activeSelection(runtime)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!state || !selection || !worksheet) return 0
  // links of the rows the grid has not streamed in yet are only known once they are loaded
  if (cellCount(selection.range) <= LOAD_BEFORE_CLEAR_MAX_CELLS) {
    await ensureRangeResident(
      runtime,
      ctx.lazyWorkbookRef,
      worksheet,
      selection.range,
      ctx.setMessage,
    )
  }
  const cells = hyperlinkCellsIn(state, selection.sheetId, selection.range)
  if (cells.length === 0) return 0
  await runAsOneUndoStepAsync(runtime, () =>
    withJournalUndo(
      ctx,
      hyperlinkKind,
      JSON.stringify({
        sheetId: selection.sheetId,
        cells: cells.map(({ row, column }) => `${row}:${column}`),
      } satisfies HyperlinkKey),
      () => {
        for (const { row, column } of cells) {
          recordHyperlinkEdit(state.editJournal, selection.sheetId, row, column, null)
          state.hyperlinkTargets.get(selection.sheetId)?.delete(`${row}:${column}`)
          if (options.removeFormat) {
            worksheet
              .getRange(row, column, 1, 1)
              .setValue({ s: { ul: null, cl: null } } as unknown as ICellData)
          }
        }
        ctx.setPendingEdits(journalSize(state.editJournal))
      },
    ),
  )
  return cells.length
}

/// Deletes the notes under the selection (Excel's Clear Comments and Notes). Returns how many went.
export async function clearNotes(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
): Promise<number> {
  const selection = activeSelection(runtime)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!selection || !worksheet) return 0
  if (cellCount(selection.range) <= LOAD_BEFORE_CLEAR_MAX_CELLS) {
    await ensureRangeResident(
      runtime,
      ctx.lazyWorkbookRef,
      worksheet,
      selection.range,
      ctx.setMessage,
    )
  }
  const { range } = selection
  const notes = worksheet
    .getNotes()
    .filter(
      (note) =>
        note.row >= range.startRow &&
        note.row <= range.endRow &&
        note.col >= range.startColumn &&
        note.col <= range.endColumn,
    )
  if (notes.length === 0) return 0
  if (ctx.lazyWorkbookRef.current) {
    const ops: WorkbookOperation[] = notes.map((note) => ({
      op: 'set_note',
      sheetId: selection.sheetId,
      address: formatAddress(note.row, note.col),
      text: null,
    }))
    await ctx.runOps(ops, null)
  } else {
    await deleteNotesUndoable(runtime, selection.sheetId, notes)
  }
  return notes.length
}

/**
 * Clear Contents (the cell's text and its hyperlink) and Clear All (also notes and formats): Univer's
 * clear plus the part it does not reach, in order, as ONE undo step — ⌘Z must not bring back the text
 * and leave the link or the note behind.
 */
export async function clearWithExtras(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  which: 'contents' | 'all',
): Promise<void> {
  await runAsOneUndoStepAsync(runtime, async () => {
    if (which === 'all') await clearNotes(ctx, runtime)
    await clearHyperlinks(ctx, runtime, { removeFormat: false })
    await runtime.univerAPI.executeCommand(
      which === 'all'
        ? 'sheet.command.clear-selection-all'
        : 'sheet.command.clear-selection-content',
    )
  })
}
