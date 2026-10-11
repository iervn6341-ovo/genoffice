import { journalSize } from './edit-journal'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'
import { rethemeStyles } from './themes'
import { loadVisibleRange } from './univer-sync'

/**
 * Page Layout → Themes / Colors / Fonts journal the theme part of the save and re-resolve the
 * styles that carry theme provenance. Undo puts the previous colors / fonts back and re-resolves.
 */
interface ThemeSnapshot {
  journalColors: { name: string; values: string[] } | null
  journalFonts: { name: string; major: string; minor: string } | null
  colors: string[] | null
  fonts: { major: string; minor: string } | null
}

const themeKind: JournalUndoKind<ThemeSnapshot> = {
  id: 'genoffice.mutation.set-document-theme',
  capture: (state) =>
    structuredClone({
      journalColors: state.editJournal.theme.colors ?? null,
      journalFonts: state.editJournal.theme.fonts ?? null,
      colors: state.file.themeColors ?? null,
      fonts: state.file.themeFonts ?? null,
    }) as ThemeSnapshot,
  restore: (state, _key, snapshot, ctx) => {
    const { theme } = state.editJournal
    if (snapshot.journalColors) theme.colors = snapshot.journalColors
    else delete theme.colors
    if (snapshot.journalFonts) theme.fonts = snapshot.journalFonts
    else delete theme.fonts
    // a theme edit only runs on a file that has both parts, so the snapshots are there
    if (snapshot.colors) state.file.themeColors = snapshot.colors
    if (snapshot.fonts) state.file.themeFonts = snapshot.fonts
    if (!state.file.themeColors) return
    state.file.styles = rethemeStyles(
      state.file.styles,
      state.file.themeColors,
      state.file.themeFonts ?? null,
    )
    for (const loadedSheetId of [...state.loadedRanges.keys()]) {
      state.loadedRanges.delete(loadedSheetId)
      state.appliedRowKeys.delete(loadedSheetId)
    }
    ctx.setPendingEdits(journalSize(state.editJournal))
    const runtime = ctx.univerRef.current
    const active = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (runtime && active)
      void loadVisibleRange(runtime, ctx.lazyWorkbookRef, active, () => undefined)
  },
}

/** Runs a theme edit and records it as one undo step when the theme changed. */
export function withThemeUndo(
  ctx: JournalUndoContext,
  edit: () => Promise<boolean | void> | boolean | void,
): Promise<void> {
  return withJournalUndo(ctx, themeKind, 'theme', edit)
}
