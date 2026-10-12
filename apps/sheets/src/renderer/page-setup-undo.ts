import { journalSize, recordPageSetup, type PageSetupJournalState } from './edit-journal'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'
import { effectivePageBreaks } from './page-break-preview'
import {
  MARGIN_PRESETS,
  resolveEffectivePageSetup,
  type EffectivePageSetup,
} from './print-settings'
import type { LazyWorkbookState } from './univer-state'

/**
 * Page Layout edits (orientation, size, margins, scaling, print area / titles, breaks, header and
 * footer) only change the journal that the save writes — see journal-undo.ts.
 *
 * The snapshot holds the sheet's journal state (null when it had no edits) AND the setup that was in
 * force before the edit. ⌘Z in the same session puts the journal state back, which is exact. ⌘Z after
 * a save reaches an empty journal — the save wrote the edit into the file and reopened it — so there
 * is no journal state to put back: the fields where the setup in force now differs from the one the
 * snapshot remembers become journal edits again, and the next save writes the older values.
 *
 * What the journal cannot say it cannot write back after a save: margins that are none of the three
 * presets (a file's own custom margins) and more than one print area stay as the save left them.
 */
interface PageSetupSnapshot {
  journal: PageSetupJournalState | null
  effective: Pick<
    EffectivePageSetup,
    | 'orientation'
    | 'paperSize'
    | 'scale'
    | 'fitToWidth'
    | 'fitToHeight'
    | 'fitToPage'
    | 'margins'
    | 'printGridlines'
    | 'printHeadings'
    | 'printAreas'
    | 'printTitles'
    | 'header'
    | 'footer'
  > & {
    /// the manual page breaks in screen coordinates; null until the sheet is indexed
    breaks?: PageBreakSet | null
  }
}

interface PageBreakSet {
  readonly rowBreaks: readonly number[]
  readonly colBreaks: readonly number[]
}

/** the page setup in force for the sheet: the file's, with this session's edits over it */
export function effectiveSetupOf(state: LazyWorkbookState, sheetId: string): EffectivePageSetup {
  const fileSheet = state.file.sheets.find((sheet) => sheet.id === sheetId)
  return resolveEffectivePageSetup(
    state.editJournal.pageSetup.get(sheetId) ?? {},
    state.sheetFilePageSetups.get(sheetId) ?? null,
    {
      ...(fileSheet?.printArea === undefined ? {} : { printArea: fileSheet.printArea }),
      ...(fileSheet?.printTitles === undefined ? {} : { printTitles: fileSheet.printTitles }),
    },
    state.editJournal.structuralOps.get(sheetId) ?? [],
  )
}

const same = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b)

/** the journal fields that bring the setup in force back to `wanted` */
export function pageSetupPatch(
  now: PageSetupSnapshot['effective'],
  wanted: PageSetupSnapshot['effective'],
): PageSetupJournalState {
  const patch: PageSetupJournalState = {}
  if (now.orientation !== wanted.orientation) patch.orientation = wanted.orientation
  if (now.paperSize !== wanted.paperSize) patch.paperSize = wanted.paperSize
  if (now.scale !== wanted.scale) patch.scale = wanted.scale
  if (now.fitToWidth !== wanted.fitToWidth) patch.fitToWidth = wanted.fitToWidth
  if (now.fitToHeight !== wanted.fitToHeight) patch.fitToHeight = wanted.fitToHeight
  if (now.fitToPage !== wanted.fitToPage) patch.fitToPage = wanted.fitToPage
  if (!same(now.margins, wanted.margins)) {
    // the journal knows the three presets; margins of another kind cannot be written back
    const preset = (Object.keys(MARGIN_PRESETS) as (keyof typeof MARGIN_PRESETS)[]).find((name) =>
      same(MARGIN_PRESETS[name], wanted.margins),
    )
    if (preset) patch.margins = preset
  }
  if (now.printGridlines !== wanted.printGridlines) patch.printGridlines = wanted.printGridlines
  if (now.printHeadings !== wanted.printHeadings) patch.printHeadings = wanted.printHeadings
  if (!same(now.printAreas, wanted.printAreas)) {
    // one print area is what the journal holds; none clears it
    if (wanted.printAreas.length <= 1) patch.printArea = wanted.printAreas[0] ?? null
  }
  if (now.printTitles !== wanted.printTitles) patch.printTitles = wanted.printTitles
  // a break set the snapshot knew, and the file now says otherwise (or has not said yet)
  if (wanted.breaks) {
    if (!now.breaks || !same(now.breaks.rowBreaks, wanted.breaks.rowBreaks)) {
      patch.rowBreaks = [...wanted.breaks.rowBreaks]
    }
    if (!now.breaks || !same(now.breaks.colBreaks, wanted.breaks.colBreaks)) {
      patch.colBreaks = [...wanted.breaks.colBreaks]
    }
  }
  if (!same(now.header, wanted.header)) patch.header = wanted.header
  if (!same(now.footer, wanted.footer)) patch.footer = wanted.footer
  return patch
}

const pickEffective = (
  setup: EffectivePageSetup,
  breaks: PageBreakSet | null,
): PageSetupSnapshot['effective'] => ({
  orientation: setup.orientation,
  paperSize: setup.paperSize,
  scale: setup.scale,
  fitToWidth: setup.fitToWidth,
  fitToHeight: setup.fitToHeight,
  fitToPage: setup.fitToPage,
  margins: setup.margins,
  printGridlines: setup.printGridlines,
  printHeadings: setup.printHeadings,
  printAreas: setup.printAreas,
  printTitles: setup.printTitles,
  header: setup.header,
  footer: setup.footer,
  breaks,
})

const pageSetupKind: JournalUndoKind<PageSetupSnapshot> = {
  id: 'genoffice.mutation.set-page-setup',
  capture: (state, sheetId) => {
    const value = state.editJournal.pageSetup.get(sheetId)
    return structuredClone({
      journal: value ? structuredClone(value) : null,
      effective: pickEffective(
        effectiveSetupOf(state, sheetId),
        effectivePageBreaks(state, sheetId),
      ),
    })
  },
  restore: (state, sheetId, snapshot, ctx) => {
    if (snapshot.journal) state.editJournal.pageSetup.set(sheetId, snapshot.journal)
    else state.editJournal.pageSetup.delete(sheetId)
    // after a save the journal cannot say what the file now lacks: write the difference back
    const patch = pageSetupPatch(
      pickEffective(effectiveSetupOf(state, sheetId), effectivePageBreaks(state, sheetId)),
      snapshot.effective,
    )
    if (Object.keys(patch).length > 0) recordPageSetup(state.editJournal, sheetId, patch)
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
