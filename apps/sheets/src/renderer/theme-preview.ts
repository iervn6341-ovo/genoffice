import type { WorkbookCellStyle } from '../shared/desktop-api'
import { COLOR_SCHEMES, FONT_SCHEMES, rethemeStyles, THEME_PRESETS } from './themes'
import { loadVisibleRange } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/**
 * Page Layout → Themes / Colors / Fonts galleries preview a choice live while the pointer is on
 * it, like Excel: the grid recolors and refonts, and moving away (or closing the gallery) puts
 * everything back. A preview changes only what the grid reads — the file's resolved styles and
 * theme — never the edit journal and never the undo stack; the real pick (page-layout.ts) starts
 * from the restored state. The snapshot is kept per workbook state.
 */
interface Snapshot {
  /// the journal's theme edits when the preview began: if they differ at the end, a pick, ⌘Z or ⇧⌘Z
  /// ran during the preview and the workbook already holds the right theme
  readonly journalTheme: string
  readonly styles: readonly WorkbookCellStyle[]
  readonly colors: readonly string[] | undefined
  readonly fonts: { major: string; minor: string } | undefined
}

const snapshots = new WeakMap<LazyWorkbookState, Snapshot>()

export type ThemePreviewKind = 'theme' | 'colors' | 'fonts'

/** the colors and fonts a gallery entry stands for, or null for an id it does not know */
export function themePreviewChoice(
  kind: ThemePreviewKind,
  id: string,
): {
  colors: readonly string[] | undefined
  fonts: { major: string; minor: string } | undefined
} | null {
  if (kind === 'theme') {
    const preset = THEME_PRESETS.find((entry) => entry.id === id)
    return preset ? { colors: preset.colors.values, fonts: preset.fonts } : null
  }
  if (kind === 'colors') {
    const scheme = COLOR_SCHEMES.find((entry) => entry.id === id)
    return scheme ? { colors: scheme.values, fonts: undefined } : null
  }
  const scheme = FONT_SCHEMES.find((entry) => entry.id === id)
  return scheme ? { colors: undefined, fonts: scheme } : null
}

function repaint(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
): void {
  const state = lazyWorkbookRef.current
  if (!state) return
  // the grid rereads its styles: drop what it has applied and reload the visible part, silently
  for (const sheetId of [...state.loadedRanges.keys()]) {
    state.loadedRanges.delete(sheetId)
    state.appliedRowKeys.delete(sheetId)
  }
  const active = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (active) void loadVisibleRange(runtime, lazyWorkbookRef, active, () => undefined)
}

export interface ThemePreviewContext {
  readonly univerRef: { readonly current: UniverRuntime | null }
  readonly lazyWorkbookRef: { current: LazyWorkbookState | null }
}

/** Shows `kind` / `id` on the grid. A workbook without a theme part (Colors / Themes) is left alone. */
export function startThemePreview(
  ctx: ThemePreviewContext,
  kind: ThemePreviewKind,
  id: string,
): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state || !state.file.themeColors) return
  if (kind === 'fonts' && !state.file.themeFonts) return
  const choice = themePreviewChoice(kind, id)
  if (!choice) return
  let base = snapshots.get(state)
  if (!base) {
    base = {
      journalTheme: JSON.stringify(state.editJournal.theme),
      styles: state.file.styles,
      colors: state.file.themeColors,
      fonts: state.file.themeFonts,
    }
    snapshots.set(state, base)
  }
  // a theme without a rewritable font scheme keeps its fonts, as the real pick does
  const fonts = state.file.themeFonts && choice.fonts ? choice.fonts : base.fonts
  const colors = choice.colors ?? base.colors
  state.file.themeColors = colors ? [...colors] : state.file.themeColors
  if (fonts) state.file.themeFonts = { major: fonts.major, minor: fonts.minor }
  state.file.styles = rethemeStyles(
    base.styles as WorkbookCellStyle[],
    state.file.themeColors,
    state.file.themeFonts ?? null,
  )
  repaint(runtime, ctx.lazyWorkbookRef)
}

/**
 * Puts the workbook's own theme back. `repaintGrid` false is for a real pick that is about to
 * repaint anyway (it must start from the restored state, not the previewed one).
 */
export function endThemePreview(ctx: ThemePreviewContext, repaintGrid = true): void {
  const state = ctx.lazyWorkbookRef.current
  if (!state) return
  const base = snapshots.get(state)
  if (!base) return
  snapshots.delete(state)
  // ⌘Z (the app menu still fires while the pointer rests on a card) rewrote the theme and
  // re-resolved the styles itself: putting the older snapshot back would undo that
  if (JSON.stringify(state.editJournal.theme) !== base.journalTheme) {
    const runtime = ctx.univerRef.current
    if (repaintGrid && runtime) repaint(runtime, ctx.lazyWorkbookRef)
    return
  }
  state.file.styles = base.styles as WorkbookCellStyle[]
  if (base.colors) state.file.themeColors = [...base.colors]
  if (base.fonts) state.file.themeFonts = { ...base.fonts }
  const runtime = ctx.univerRef.current
  if (repaintGrid && runtime) repaint(runtime, ctx.lazyWorkbookRef)
}

export function isThemePreviewing(state: LazyWorkbookState | null): boolean {
  return state !== null && snapshots.has(state)
}
