import type { WorkbookStyleEdit } from '../shared/desktop-api'
import { journalSize, recordNeutralStyleEdit, type JournalEntry } from './edit-journal'
import {
  registerJournalUndoKind,
  withJournalUndo,
  type JournalUndoContext,
  type JournalUndoKind,
} from './journal-undo'
import type { CellProtectionFlags, SelectionProtection } from './selection-protection'
import type { LazyWorkbookState } from './univer-state'
import { borderDeltaForCell, type BorderAction } from './border-format'

export interface ProtectionSnapshot {
  readonly sessionId: string
  readonly cells: readonly {
    row: number
    column: number
    styleReference?: number | null
    locked: boolean
    hidden: boolean
    horizontalAlignment?: string
    verticalAlignment?: string
    textRotation?: number
    shrinkToFit?: boolean
    underlineStyle?: CellProtectionFlags['underlineStyle']
    fontVerticalAlign?: CellProtectionFlags['fontVerticalAlign']
    entry: JournalEntry | null
  }[]
}

export function restoreProtectionSnapshot(
  state: Pick<LazyWorkbookState, 'file' | 'editJournal'>,
  sheetId: string,
  snapshot: ProtectionSnapshot,
  companion?: ProtectionSnapshot,
): void {
  if (state.file.sessionId === snapshot.sessionId) {
    let entries = state.editJournal.cells.get(sheetId)
    if (!entries) {
      entries = new Map()
      state.editJournal.cells.set(sheetId, entries)
    }
    for (const cell of snapshot.cells) {
      const cellKey = `${cell.row}:${cell.column}`
      if (cell.entry) entries.set(cellKey, cell.entry)
      else entries.delete(cellKey)
    }
    if (entries.size === 0) state.editJournal.cells.delete(sheetId)
  } else {
    // A save moved the flags into the base xf; restore only the flags this
    // history step changed, leaving cell content and other styles intact.
    for (const [index, cell] of snapshot.cells.entries()) {
      const opposite = companion?.cells[index]
      if (!opposite) continue
      if (cell.styleReference !== undefined) {
        let entries = state.editJournal.cells.get(sheetId)
        if (!entries) {
          entries = new Map()
          state.editJournal.cells.set(sheetId, entries)
        }
        const key = `${cell.row}:${cell.column}`
        const current = entries.get(key)
        const priorStyle = cell.entry?.style
        const priorReset = cell.entry?.styleReset === true
        const content = current
          ? {
              row: current.row,
              column: current.column,
              hasValue: current.hasValue,
              value: current.value,
              ...(current.formula === undefined ? {} : { formula: current.formula }),
              ...(current.rich === undefined ? {} : { rich: current.rich }),
            }
          : { row: cell.row, column: cell.column, hasValue: false, value: null }
        entries.set(key, {
          ...content,
          ...(priorStyle === undefined ? {} : { style: priorStyle }),
          ...(priorReset ? { styleReset: true } : { styleReference: cell.styleReference }),
        })
        continue
      }
      const delta: WorkbookStyleEdit = {}
      if (cell.locked !== opposite.locked) delta.protectionLocked = cell.locked
      if (cell.hidden !== opposite.hidden) delta.protectionHidden = cell.hidden
      if (cell.horizontalAlignment !== opposite.horizontalAlignment && cell.horizontalAlignment) {
        delta.horizontalAlignment =
          cell.horizontalAlignment as WorkbookStyleEdit['horizontalAlignment']
      }
      if (cell.verticalAlignment !== opposite.verticalAlignment && cell.verticalAlignment) {
        delta.verticalAlignment = cell.verticalAlignment as WorkbookStyleEdit['verticalAlignment']
      }
      if (cell.textRotation !== opposite.textRotation && cell.textRotation !== undefined) {
        delta.textRotation = cell.textRotation
      }
      if (cell.shrinkToFit !== opposite.shrinkToFit && cell.shrinkToFit !== undefined) {
        delta.shrinkToFit = cell.shrinkToFit
      }
      if (cell.underlineStyle !== opposite.underlineStyle && cell.underlineStyle !== undefined) {
        delta.underline = cell.underlineStyle !== 'none'
        if (cell.underlineStyle !== 'none') delta.underlineStyle = cell.underlineStyle
      }
      if (
        cell.fontVerticalAlign !== opposite.fontVerticalAlign &&
        cell.fontVerticalAlign !== undefined
      ) {
        delta.fontVerticalAlign = cell.fontVerticalAlign
      }
      if (Object.keys(delta).length > 0) {
        recordNeutralStyleEdit(state.editJournal, sheetId, cell.row, cell.column, delta)
      }
    }
  }
}

const protectionKind: JournalUndoKind<ProtectionSnapshot> = {
  id: 'genoffice.mutation.set-cell-protection',
  capture: (state, key) => {
    const { sheetId, selection } = JSON.parse(key) as {
      sheetId: string
      selection: SelectionProtection
    }
    const entries = state.editJournal.cells.get(sheetId)
    return {
      sessionId: state.file.sessionId,
      cells: selection.cells.map((cell) => {
        const entry = entries?.get(`${cell.row}:${cell.column}`) ?? null
        return {
          row: cell.row,
          column: cell.column,
          ...(cell.styleReference === undefined ? {} : { styleReference: cell.styleReference }),
          locked: entry?.style?.protectionLocked ?? cell.locked,
          hidden: entry?.style?.protectionHidden ?? cell.hidden,
          horizontalAlignment: entry?.style?.horizontalAlignment ?? cell.horizontalAlignment,
          verticalAlignment: entry?.style?.verticalAlignment ?? cell.verticalAlignment,
          textRotation: entry?.style?.textRotation ?? cell.textRotation,
          shrinkToFit: entry?.style?.shrinkToFit ?? cell.shrinkToFit,
          underlineStyle:
            entry?.style?.underline === false
              ? 'none'
              : (entry?.style?.underlineStyle ?? cell.underlineStyle),
          fontVerticalAlign: entry?.style?.fontVerticalAlign ?? cell.fontVerticalAlign,
          entry: entry ? structuredClone(entry) : null,
        }
      }),
    }
  },
  restore: (state, key, snapshot, ctx, companion) => {
    const { sheetId } = JSON.parse(key) as { sheetId: string }
    restoreProtectionSnapshot(state, sheetId, snapshot, companion)
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.refreshAlignmentPreview?.(sheetId)
  },
}

export function registerCellProtectionUndoRuntime(ctx: JournalUndoContext): void {
  registerJournalUndoKind(ctx, protectionKind)
}

export function protectionDeltaOfCommand(command: string): WorkbookStyleEdit | null {
  const match = /^cellprot:(locked-(?:on|off))?:(hidden-(?:on|off))?$/.exec(command)
  if (!match) return null
  const delta: WorkbookStyleEdit = {}
  if (match[1]) delta.protectionLocked = match[1] === 'locked-on'
  if (match[2]) delta.protectionHidden = match[2] === 'hidden-on'
  return Object.keys(delta).length > 0 ? delta : null
}

export function alignmentDeltaOfCommand(command: string): WorkbookStyleEdit | null {
  const [kind, value] = command.split(':')
  if (
    kind === 'align' &&
    value &&
    [
      'general',
      'left',
      'center',
      'right',
      'fill',
      'justify',
      'centerContinuous',
      'distributed',
    ].includes(value)
  ) {
    return { horizontalAlignment: value as WorkbookStyleEdit['horizontalAlignment'] }
  }
  if (
    kind === 'valign' &&
    value &&
    ['top', 'middle', 'bottom', 'justify', 'distributed'].includes(value)
  ) {
    return {
      verticalAlignment: (value === 'middle'
        ? 'center'
        : value) as WorkbookStyleEdit['verticalAlignment'],
    }
  }
  if (kind === 'rotate-ooxml' && value) {
    const angle = Number(value)
    if (Number.isInteger(angle) && ((angle >= 0 && angle <= 180) || angle === 255))
      return { textRotation: angle }
  }
  if (kind === 'shrink' && (value === 'on' || value === 'off')) {
    return { shrinkToFit: value === 'on' }
  }
  return null
}

export function fontEffectDeltaOfCommand(command: string): WorkbookStyleEdit | null {
  const match =
    /^font-effect:(none|single|double|singleAccounting|doubleAccounting)?:(baseline|superscript|subscript)?$/.exec(
      command,
    )
  if (!match) return null
  const delta: WorkbookStyleEdit = {}
  if (match[1]) {
    delta.underline = match[1] !== 'none'
    if (match[1] !== 'none')
      delta.underlineStyle = match[1] as NonNullable<WorkbookStyleEdit['underlineStyle']>
  }
  if (match[2])
    delta.fontVerticalAlign = match[2] as NonNullable<WorkbookStyleEdit['fontVerticalAlign']>
  return Object.keys(delta).length > 0 ? delta : null
}

export function withCellProtectionUndo(
  ctx: JournalUndoContext,
  selection: SelectionProtection,
  delta: WorkbookStyleEdit,
): Promise<void> {
  const key = JSON.stringify({ sheetId: selection.sheetId, selection })
  return withJournalUndo(ctx, protectionKind, key, () => {
    const state = ctx.lazyWorkbookRef.current
    if (!state) return false
    for (const cell of selection.cells) {
      recordNeutralStyleEdit(state.editJournal, selection.sheetId, cell.row, cell.column, delta)
    }
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.refreshAlignmentPreview?.(selection.sheetId)
  })
}

export function withSelectionBorderUndo(
  ctx: JournalUndoContext,
  selection: SelectionProtection,
  action: BorderAction,
): Promise<void> {
  const key = JSON.stringify({ sheetId: selection.sheetId, selection })
  return withJournalUndo(ctx, protectionKind, key, () => {
    const state = ctx.lazyWorkbookRef.current
    if (!state) return false
    let changed = false
    for (const cell of selection.cells) {
      const delta = borderDeltaForCell(action, selection.range, cell, selection.merges)
      if (!delta) continue
      recordNeutralStyleEdit(state.editJournal, selection.sheetId, cell.row, cell.column, delta)
      changed = true
    }
    if (!changed) return false
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.refreshAlignmentPreview?.(selection.sheetId)
  })
}
