import { journalSize, recordPageSetup, type PageSetupJournalState } from './edit-journal'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'
import { effectivePageBreaks } from './page-break-preview'
import {
  MARGIN_PRESETS,
  resolveHeaderFooterEditVariants,
  resolveEffectivePageSetup,
  type EffectivePageSetup,
  type HeaderFooterEditVariant,
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
    firstPage: HeaderFooterEditVariant
    evenPages: HeaderFooterEditVariant
    headerFooterPictures?: Array<{
      id: string
      position: string
      widthPt: number
      heightPt: number
      mediaType: string
      base64?: string
    }>
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

function pictureMime(
  value: string,
): value is NonNullable<
  NonNullable<NonNullable<PageSetupJournalState['headerFooterPictureEdits']>[number]['image']>
>['mediaType'] {
  return [
    'image/png',
    'image/jpeg',
    'image/gif',
    'image/x-emf',
    'image/x-wmf',
    'image/bmp',
    'image/webp',
    'image/svg+xml',
    'image/x-emz',
    'image/x-wmz',
  ].includes(value)
}

function storedVariantsOf(state: LazyWorkbookState, sheetId: string) {
  return resolveHeaderFooterEditVariants(
    state.editJournal.pageSetup.get(sheetId) ?? {},
    state.sheetFilePageSetups.get(sheetId) ?? null,
  )
}

/** the journal fields that bring the setup in force back to `wanted` */
export function pageSetupPatch(
  now: PageSetupSnapshot['effective'],
  wanted: PageSetupSnapshot['effective'],
  picturePositions?: ReadonlySet<string>,
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
  const restoreVariant = (key: 'firstPage' | 'evenPages'): void => {
    const current = now[key]
    const target = wanted[key]
    if (same(current, target)) return
    if (target.enabled) {
      patch[key] = target.parts ?? { header: null, footer: null }
      return
    }
    if (current.enabled) patch[key] = null
    if (!same(current.parts, target.parts)) {
      patch[key === 'firstPage' ? 'firstPageStored' : 'evenPagesStored'] = target.parts
    }
  }
  restoreVariant('firstPage')
  restoreVariant('evenPages')
  const currentPictures = new Map(
    (now.headerFooterPictures ?? []).map((picture) => [picture.position, picture]),
  )
  const wantedPictures = new Map(
    (wanted.headerFooterPictures ?? []).map((picture) => [picture.position, picture]),
  )
  const pictureEdits: NonNullable<PageSetupJournalState['headerFooterPictureEdits']> = []
  for (const position of new Set([...currentPictures.keys(), ...wantedPictures.keys()])) {
    if (picturePositions && !picturePositions.has(position)) continue
    const current = currentPictures.get(position)
    const target = wantedPictures.get(position)
    if (same(current, target)) continue
    if (!target) {
      pictureEdits.push({ position, image: null })
    } else if (target.base64 !== undefined && pictureMime(target.mediaType)) {
      pictureEdits.push({
        position,
        image: { mediaType: target.mediaType, base64: target.base64 },
        widthPt: target.widthPt,
        heightPt: target.heightPt,
      })
    } else if (current?.id === target.id) {
      pictureEdits.push({ position, widthPt: target.widthPt, heightPt: target.heightPt })
    }
  }
  if (pictureEdits.length > 0) patch.headerFooterPictureEdits = pictureEdits
  return patch
}

const pickEffective = (
  state: LazyWorkbookState,
  sheetId: string,
  setup: EffectivePageSetup,
  breaks: PageBreakSet | null,
  variants: ReturnType<typeof storedVariantsOf>,
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
  firstPage: variants.firstPage,
  evenPages: variants.evenPages,
  headerFooterPictures: setup.headerFooterPictures.map((picture) => {
    const pending = state.editJournal.pageSetup
      .get(sheetId)
      ?.headerFooterPictureEdits?.find((edit) => edit.position === picture.position)?.image
    const source =
      pending && pending !== null
        ? pending
        : state.headerFooterPictureData.get(sheetId)?.get(picture.position)
    return {
      ...picture,
      ...(source ? { base64: source.base64 } : {}),
    }
  }),
  breaks,
})

const pageSetupKind: JournalUndoKind<PageSetupSnapshot> = {
  id: 'genoffice.mutation.set-page-setup',
  capture: (state, sheetId) => {
    const value = state.editJournal.pageSetup.get(sheetId)
    return structuredClone({
      journal: value ? structuredClone(value) : null,
      effective: pickEffective(
        state,
        sheetId,
        effectiveSetupOf(state, sheetId),
        effectivePageBreaks(state, sheetId),
        storedVariantsOf(state, sheetId),
      ),
    })
  },
  restore: (state, sheetId, snapshot, ctx, companion) => {
    if (snapshot.journal) state.editJournal.pageSetup.set(sheetId, snapshot.journal)
    else state.editJournal.pageSetup.delete(sheetId)
    // after a save the journal cannot say what the file now lacks: write the difference back
    const picturePositions = new Set<string>()
    if (companion) {
      const before = new Map(
        (snapshot.effective.headerFooterPictures ?? []).map((picture) => [
          picture.position,
          picture,
        ]),
      )
      const after = new Map(
        (companion.effective.headerFooterPictures ?? []).map((picture) => [
          picture.position,
          picture,
        ]),
      )
      for (const position of new Set([...before.keys(), ...after.keys()])) {
        if (!same(before.get(position), after.get(position))) picturePositions.add(position)
      }
    }
    const patch = pageSetupPatch(
      pickEffective(
        state,
        sheetId,
        effectiveSetupOf(state, sheetId),
        effectivePageBreaks(state, sheetId),
        storedVariantsOf(state, sheetId),
      ),
      snapshot.effective,
      picturePositions,
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
