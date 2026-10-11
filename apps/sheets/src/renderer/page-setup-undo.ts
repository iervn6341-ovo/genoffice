import { journalSize, type PageSetupJournalState } from './edit-journal'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'

/**
 * Page Layout edits (orientation, size, margins, scaling, print area / titles, breaks, header and
 * footer) only change the journal that the save writes — see journal-undo.ts. The snapshot is the
 * sheet's page-setup state, null when it had no edits.
 */
const pageSetupKind: JournalUndoKind<PageSetupJournalState | null> = {
  id: 'genoffice.mutation.set-page-setup',
  capture: (state, sheetId) => {
    const value = state.editJournal.pageSetup.get(sheetId)
    return value ? structuredClone(value) : null
  },
  restore: (state, sheetId, snapshot, ctx) => {
    if (snapshot) state.editJournal.pageSetup.set(sheetId, snapshot)
    else state.editJournal.pageSetup.delete(sheetId)
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.refreshPageBreakPreview?.()
  },
}

export type PageSetupUndoContext = JournalUndoContext

/** Runs a page-setup edit and records it as one undo step when the sheet's state changed. */
export function withPageSetupUndo(
  ctx: PageSetupUndoContext,
  sheetId: string,
  edit: () => Promise<boolean | void> | boolean | void,
): Promise<void> {
  return withJournalUndo(ctx, pageSetupKind, sheetId, edit)
}
