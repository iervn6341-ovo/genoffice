import { journalSize, recordThemeColors, recordThemeFonts } from './edit-journal'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'
import { COLOR_SCHEMES, FONT_SCHEMES, rethemeStyles } from './themes'
import { loadVisibleRange } from './univer-sync'
import type { LazyWorkbookState } from './univer-state'

/**
 * Page Layout → Themes / Colors / Fonts journal the theme part of the save and re-resolve the
 * styles that carry theme provenance. Undo puts the previous colors / fonts back and re-resolves.
 *
 * What the file on disk holds is the baseline the journal is a difference from. In the session that
 * made the edit it is the palette before the first edit (remembered here); in a session reopened by a
 * save it is the palette the file loaded with. ⌘Z that reaches a palette other than the baseline
 * writes it into the journal, so the next save puts it back in the file — a save in between made
 * the edit part of the file, and an empty journal would leave it there.
 */
interface ThemeSnapshot {
  journalColors: { name: string; values: string[] } | null
  journalFonts: { name: string; major: string; minor: string } | null
  colors: string[] | null
  fonts: { major: string; minor: string } | null
}

type Baseline = NonNullable<LazyWorkbookState['themeBaseline']>

/** the palette and fonts the file on disk holds, as far as this session knows them */
function baselineOf(state: LazyWorkbookState): Baseline {
  const { theme } = state.editJournal
  const known = state.themeBaseline
  return {
    // an edit in the journal means the live palette is not the file's: only a remembered baseline knows it
    colors: theme.colors ? (known?.colors ?? null) : (state.file.themeColors ?? null),
    fonts: theme.fonts ? (known?.fonts ?? null) : (state.file.themeFonts ?? null),
  }
}

const sameColors = (a: readonly string[] | null, b: readonly string[] | null): boolean =>
  a !== null && b !== null && a.length === b.length && a.every((v, i) => v === b[i])

const sameFonts = (
  a: { major: string; minor: string } | null,
  b: { major: string; minor: string } | null,
): boolean => a !== null && b !== null && a.major === b.major && a.minor === b.minor

const themeKind: JournalUndoKind<ThemeSnapshot> = {
  id: 'genoffice.mutation.set-document-theme',
  capture: (state) => {
    // the first edit of a session knows the file's theme: it is the live one
    state.themeBaseline ??= baselineOf(state)
    return structuredClone({
      journalColors: state.editJournal.theme.colors ?? null,
      journalFonts: state.editJournal.theme.fonts ?? null,
      colors: state.file.themeColors ?? null,
      fonts: state.file.themeFonts ?? null,
    }) as ThemeSnapshot
  },
  restore: (state, _key, snapshot, ctx) => {
    const { theme } = state.editJournal
    // what the file holds, found before the snapshot overwrites the live palette
    state.themeBaseline ??= baselineOf(state)
    const baseline = state.themeBaseline
    if (snapshot.journalColors) theme.colors = snapshot.journalColors
    else delete theme.colors
    if (snapshot.journalFonts) theme.fonts = snapshot.journalFonts
    else delete theme.fonts
    // a theme edit only runs on a file that has both parts, so the snapshots are there
    if (snapshot.colors) state.file.themeColors = snapshot.colors
    if (snapshot.fonts) state.file.themeFonts = snapshot.fonts
    // the palette the snapshot calls for is not the one in the file: the save has to write it
    if (
      !snapshot.journalColors &&
      snapshot.colors &&
      !sameColors(snapshot.colors, baseline.colors)
    ) {
      recordThemeColors(
        state.editJournal,
        COLOR_SCHEMES.find((scheme) => sameColors(scheme.values, snapshot.colors))?.name ??
          'Office',
        snapshot.colors,
      )
    }
    if (!snapshot.journalFonts && snapshot.fonts && !sameFonts(snapshot.fonts, baseline.fonts)) {
      const fonts = snapshot.fonts
      recordThemeFonts(
        state.editJournal,
        FONT_SCHEMES.find((scheme) => sameFonts(scheme, fonts))?.name ?? 'Office',
        fonts.major,
        fonts.minor,
      )
    }
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
