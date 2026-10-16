/**
 * Page Layout commands, header/footer, freeze journaling and PDF export.
 * Extracted from App.tsx; the App component passes a PageLayoutContext built
 * fresh per call so refs and state never go stale. Page-setup edits journal
 * per-sheet print settings; nothing renders in the grid (Univer has no
 * page-layout view), everything lands in the saved file.
 */
import { isMetafileMime, metafileToDataUrl } from '@genoffice/docx-engine/metafile'
import type { WorkbookExportPdfRequest, WorkbookPrintJob } from '../shared/desktop-api'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { ApplyOutcome } from '@genoffice/xlsx-gateway/domain/workbook.types'

import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  isSheetRemoved,
  journalSize,
  recordPageSetup,
  recordThemeColors,
  recordThemeFonts,
  type PageSetupJournalState,
} from './edit-journal'
import type { HeaderFooterResult } from './HeaderFooterDialog'
import { t } from './i18n/locale'
import { effectivePageBreaks } from './page-break-preview'
import { withPageSetupUndo } from './page-setup-undo'
import { endThemePreview, startThemePreview, type ThemePreviewKind } from './theme-preview'
import { withThemeUndo } from './theme-undo'
import { COLOR_SCHEMES, FONT_SCHEMES, rethemeStyles, THEME_PRESETS } from './themes'
import { loadVisibleRange } from './univer-sync'
import { ensureRangeResident } from './lazy-resident'
import { PRINT_JOB_MAX_CELLS, totalCells, unitsForAreas, type PrintArea } from './print-bands'
import { parseArea } from './print-html'
import {
  buildSheetPrintPayload,
  type HeaderFooterPictureImage,
  type PrintWorksheet,
} from './print-html'
import type { PrintOnlyOptions } from './print-options'
import { resolveEffectivePageSetup, type HeaderFooterPictureSlot } from './print-settings'
import { settleVisualNodes, snapshotPrintVisuals } from './print-visuals'
import { installedVisualFrames, type InstalledVisualFrame } from './WorkbookVisuals'
import { lazySheetScreenExtent, type LazyWorkbookState, type UniverRuntime } from './univer-state'

const PAPER_NAMES: Record<string, string> = {
  1: 'Letter',
  3: 'Tabloid',
  5: 'Legal',
  7: 'Executive',
  8: 'A3',
  9: 'A4',
  11: 'A5',
  12: 'B4 (JIS)',
  13: 'B5 (JIS)',
}

/** The App refs/state the page-layout actions need; built fresh per call. */
export interface PageLayoutContext {
  univerRef: { readonly current: UniverRuntime | null }
  /// The App's live ref (not a snapshot): loadVisibleRange's staleness
  /// guards compare against `.current` after awaits.
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  /// Re-renders the Page Break Preview overlay when page geometry changed.
  refreshPageBreakPreview?: () => void
  /// Re-renders the ribbon after a page-setup undo / redo changed what it shows.
  requestRender?: () => void
  /// Re-queues the floating visuals' install so a print right after load
  /// (headless export) finds their frames; optional for callers without visuals.
  requestVisualInstall?: () => void
  /// Page-setup edits run as set_page_setup ops through the shared executor.
  runOps: (
    ops: readonly WorkbookOperation[],
    successMessage?: string | null,
  ) => Promise<ApplyOutcome>
}

const PAGE_SETUP_OP_FIELDS = new Set([
  'orientation',
  'paperSize',
  'scale',
  'fitToWidth',
  'fitToHeight',
  'fitToPage',
  'margins',
  'printGridlines',
  'printHeadings',
  'printArea',
])

/// The set_page_setup op for the fields it carries (fitToPage is derived by the executor).
function pageSetupOp(sheetId: string, patch: PageSetupJournalState): WorkbookOperation {
  return {
    op: 'set_page_setup',
    sheetId,
    ...(patch.orientation !== undefined ? { orientation: patch.orientation } : {}),
    ...(patch.paperSize !== undefined ? { paperSize: patch.paperSize } : {}),
    ...(patch.scale !== undefined ? { scale: patch.scale } : {}),
    ...(patch.fitToWidth !== undefined ? { fitToWidth: patch.fitToWidth } : {}),
    ...(patch.fitToHeight !== undefined ? { fitToHeight: patch.fitToHeight } : {}),
    ...(patch.margins !== undefined ? { margins: patch.margins } : {}),
    ...(patch.printGridlines !== undefined ? { printGridlines: patch.printGridlines } : {}),
    ...(patch.printHeadings !== undefined ? { printHeadings: patch.printHeadings } : {}),
    ...(patch.printArea !== undefined ? { printArea: patch.printArea } : {}),
  }
}

/// The print pane's page-setup edits (orientation, paper, margins, scaling, gridlines,
/// headings) as one set_page_setup op on the active sheet, resolving once it applied.
export async function applyPageSetupPatch(
  ctx: PageLayoutContext,
  patch: PageSetupJournalState,
): Promise<boolean> {
  const runtime = ctx.univerRef.current
  const sheetId = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
  const state = ctx.lazyWorkbookRef.current
  if (!sheetId || !state || isSheetRemoved(state.editJournal, sheetId)) return false
  let applied = false
  await withPageSetupUndo(ctx, sheetId, async () => {
    applied = (await ctx.runOps([pageSetupOp(sheetId, patch)], null)).ok
    return applied
  })
  return applied
}

export function handlePageLayoutCommand(ctx: PageLayoutContext, rest: string): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime) return
  if (!state) {
    ctx.setMessage(t('appPageSetupNeedsFile'))
    return
  }
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const sheetId = worksheet?.getSheetId()
  if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) return
  // Themes / Colors / Fonts galleries show a choice live while the pointer is on it
  if (rest === 'theme-preview-end') {
    endThemePreview(ctx)
    return
  }
  if (rest.startsWith('theme-preview:')) {
    const [kind, id] = rest.slice('theme-preview:'.length).split(':')
    if (kind === 'theme' || kind === 'colors' || kind === 'fonts') {
      startThemePreview(ctx, kind as ThemePreviewKind, id ?? '')
    }
    return
  }
  const recordDirect = (patch: PageSetupJournalState, note: string): void => {
    void withPageSetupUndo(ctx, sheetId, () => {
      recordPageSetup(state.editJournal, sheetId, patch)
      ctx.setPendingEdits(journalSize(state.editJournal))
      ctx.setMessage(t('appPageSetupRecorded', { note }))
      ctx.refreshPageBreakPreview?.()
    })
  }
  // Fields set_page_setup carries (fitToPage is derived by the executor);
  // breaks and print titles have no op yet and journal directly.
  const record = (patch: PageSetupJournalState, note: string): void => {
    if (!Object.keys(patch).every((key) => PAGE_SETUP_OP_FIELDS.has(key))) {
      recordDirect(patch, note)
      return
    }
    const op = pageSetupOp(sheetId, patch)
    void withPageSetupUndo(ctx, sheetId, async () => {
      const outcome = await ctx.runOps([op], t('appPageSetupRecorded', { note }))
      if (outcome.ok) ctx.refreshPageBreakPreview?.()
      return outcome.ok
    })
  }
  const separator = rest.indexOf(':')
  const key = separator === -1 ? rest : rest.slice(0, separator)
  const value = separator === -1 ? '' : rest.slice(separator + 1)
  const prior = state.editJournal.pageSetup.get(sheetId) ?? {}
  switch (key) {
    case 'orientation':
      if (value !== 'portrait' && value !== 'landscape') return
      record(
        { orientation: value },
        value === 'portrait' ? t('appOrientationPortrait') : t('appOrientationLandscape'),
      )
      return
    case 'margins':
      if (value !== 'normal' && value !== 'wide' && value !== 'narrow') return
      record(
        { margins: value },
        t(
          value === 'normal'
            ? 'appMarginsNormal'
            : value === 'wide'
              ? 'appMarginsWide'
              : 'appMarginsNarrow',
        ),
      )
      return
    case 'paper': {
      const code = Number(value)
      if (!Number.isInteger(code) || code < 1 || code > 118) return
      record({ paperSize: code }, t('appPaperSizeNote', { name: PAPER_NAMES[value] ?? value }))
      return
    }
    case 'scale': {
      const scale = Number(value)
      if (!Number.isInteger(scale) || scale < 10 || scale > 400) return
      record({ scale, fitToPage: false }, t('appPrintScaleNote', { scale }))
      return
    }
    case 'fit-width':
    case 'fit-height': {
      const pages = Number(value)
      if (!Number.isInteger(pages) || pages < 0 || pages > 1_000) return
      // Excel treats the untouched other axis as Automatic (0) once
      // fit-to-page engages.
      const fitToWidth = key === 'fit-width' ? pages : (prior.fitToWidth ?? 0)
      const fitToHeight = key === 'fit-height' ? pages : (prior.fitToHeight ?? 0)
      const fitValue = pages === 0 ? t('appFitAutomatic') : t('appFitPages', { count: pages })
      record(
        { fitToWidth, fitToHeight, fitToPage: fitToWidth > 0 || fitToHeight > 0 },
        key === 'fit-width'
          ? t('appFitWidthNote', { value: fitValue })
          : t('appFitHeightNote', { value: fitValue }),
      )
      return
    }
    case 'print-gridlines':
      record(
        { printGridlines: value === '1' },
        value === '1' ? t('appGridlinesWillPrint') : t('appGridlinesWontPrint'),
      )
      return
    case 'print-headings':
      record(
        { printHeadings: value === '1' },
        value === '1' ? t('appHeadingsWillPrint') : t('appHeadingsWontPrint'),
      )
      return
    case 'print-area': {
      if (value === 'clear') {
        record({ printArea: null }, t('appPrintAreaCleared'))
        return
      }
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range) {
        ctx.setMessage(t('appSelectPrintRange'))
        return
      }
      const startColumn = range.getColumn()
      const startRow = range.getRow()
      const area =
        `${columnLabel(startColumn)}${startRow + 1}` +
        `:${columnLabel(startColumn + range.getWidth() - 1)}${startRow + range.getHeight()}`
      record({ printArea: area }, t('appPrintAreaNote', { area }))
      return
    }
    case 'theme':
    case 'theme-colors':
    case 'theme-fonts': {
      // a real pick starts from the workbook's own theme, not from the one being previewed
      endThemePreview(ctx, false)
      if (!state.file.themeColors || (key === 'theme-fonts' && !state.file.themeFonts)) {
        ctx.setMessage(t('appThemeNeedsThemePart'))
        return
      }
      // a theme's colors are its own (Excel's built-in themes are not entries of the Colors gallery)
      const colors =
        key === 'theme-fonts'
          ? undefined
          : key === 'theme'
            ? THEME_PRESETS.find((entry) => entry.id === value)?.colors
            : COLOR_SCHEMES.find((entry) => entry.id === value)
      // A theme without a rewritable fontScheme keeps its fonts: journaling
      // one would poison every subsequent save.
      const fonts =
        key === 'theme-colors' || !state.file.themeFonts
          ? undefined
          : key === 'theme'
            ? THEME_PRESETS.find((entry) => entry.id === value)?.fonts
            : FONT_SCHEMES.find((entry) => entry.id === value)
      if (colors === undefined && fonts === undefined) return
      const fileColors = state.file.themeColors
      void withThemeUndo(ctx, () => {
        if (colors !== undefined) {
          recordThemeColors(state.editJournal, colors.name, colors.values)
          state.file.themeColors = [...colors.values]
        }
        if (fonts !== undefined) {
          recordThemeFonts(state.editJournal, fonts.name, fonts.major, fonts.minor)
          state.file.themeFonts = { major: fonts.major, minor: fonts.minor }
        }
        // Live re-resolution: styles carrying theme provenance recolor now;
        // charts/CF colors follow on the save's reopen (which reparses the
        // rewritten theme part).
        state.file.styles = rethemeStyles(
          state.file.styles,
          state.file.themeColors ?? fileColors,
          state.file.themeFonts ?? null,
        )
        for (const loadedSheetId of [...state.loadedRanges.keys()]) {
          state.loadedRanges.delete(loadedSheetId)
          state.appliedRowKeys.delete(loadedSheetId)
        }
        ctx.setPendingEdits(journalSize(state.editJournal))
        const active = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
        if (active) {
          // Silent repaint: this reload only re-applies the re-themed styles;
          // its async streaming progress would otherwise land after (and wipe)
          // the theme confirmation below.
          void loadVisibleRange(runtime, ctx.lazyWorkbookRef, active, () => undefined)
        }
      })
      ctx.setMessage(t('appThemeApplied', { name: (colors ?? fonts)?.name ?? value }))
      return
    }
    case 'breaks': {
      if (value !== 'insert' && value !== 'remove' && value !== 'reset') return
      const current = effectivePageBreaks(state, sheetId)
      if (current === null) {
        ctx.setMessage(t('appBreaksNeedIndexed'))
        return
      }
      if (value === 'reset') {
        record({ rowBreaks: [], colBreaks: [] }, t('appBreaksReset'))
        return
      }
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range || !worksheet) {
        ctx.setMessage(t('appBreaksNeedCell'))
        return
      }
      const row = range.getRow()
      const column = range.getColumn()
      // Excel: a full-row selection places only the horizontal break, a
      // full-column selection only the vertical one; a cell places both.
      const fullRow = range.getWidth() >= worksheet.getMaxColumns()
      const fullColumn = range.getHeight() >= worksheet.getMaxRows()
      const wantRow = !fullColumn && row > 0
      const wantColumn = !fullRow && column > 0
      if (value === 'insert') {
        if (!wantRow && !wantColumn) {
          ctx.setMessage(t('appBreaksNeedCell'))
          return
        }
        const rowBreaks = wantRow
          ? [...new Set([...current.rowBreaks, row])].sort((a, b) => a - b)
          : current.rowBreaks
        const colBreaks = wantColumn
          ? [...new Set([...current.colBreaks, column])].sort((a, b) => a - b)
          : current.colBreaks
        record({ rowBreaks, colBreaks }, t('appBreakInserted'))
        return
      }
      const rowBreaks = current.rowBreaks.filter((id) => !(wantRow && id === row))
      const colBreaks = current.colBreaks.filter((id) => !(wantColumn && id === column))
      if (
        rowBreaks.length === current.rowBreaks.length &&
        colBreaks.length === current.colBreaks.length
      ) {
        ctx.setMessage(t('appBreakNoneHere'))
        return
      }
      record({ rowBreaks, colBreaks }, t('appBreakRemoved'))
      return
    }
    case 'print-titles': {
      if (value === 'clear') {
        record({ printTitles: null }, t('appPrintTitlesCleared'))
        return
      }
      if (value === 'first-row') {
        record({ printTitles: '1:1' }, t('appRow1Repeats'))
        return
      }
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range) {
        ctx.setMessage(t('appSelectRepeatRows'))
        return
      }
      const rows = `${range.getRow() + 1}:${range.getRow() + range.getHeight()}`
      record({ printTitles: rows }, t('appRowsRepeat', { rows }))
      return
    }
    default:
      return
  }
}

/// Insert → Header & Footer OK: journals the printed header/footer of the
/// active sheet; the save writes the worksheet's <headerFooter> element.
export function handleApplyHeaderFooter(
  ctx: PageLayoutContext,
  result: HeaderFooterResult,
): string | null {
  const state = ctx.lazyWorkbookRef.current
  if (!state) return t('appHfNeedsFile')
  const sheetId = ctx.univerRef.current?.univerAPI
    .getActiveWorkbook()
    ?.getActiveSheet()
    ?.getSheetId()
  if (!sheetId || result.sheetId !== sheetId || isSheetRemoved(state.editJournal, sheetId)) {
    return t('appActiveSheetUnavailable')
  }
  const changedPictures = new Set(result.pictureEdits.map((edit) => edit.position))
  state.headerFooterPictureData.set(
    sheetId,
    new Map(
      result.loadedPictures
        .filter((picture) => changedPictures.has(picture.position))
        .map((picture) => [
          picture.position,
          {
            mediaType: picture.mediaType,
            base64: picture.base64,
          },
        ]),
    ),
  )
  void withPageSetupUndo(ctx, sheetId, () => {
    const previous = state.editJournal.pageSetup.get(sheetId)?.headerFooterPictureEdits ?? []
    const editsByPosition = new Map(previous.map((edit) => [edit.position, edit]))
    for (const edit of result.pictureEdits) editsByPosition.set(edit.position, edit)
    recordPageSetup(state.editJournal, sheetId, {
      header: result.header,
      footer: result.footer,
      firstPage: result.firstPage,
      evenPages: result.evenPages,
      firstPageStored: result.firstPageStored,
      evenPagesStored: result.evenPagesStored,
      ...(result.pictureEdits.length > 0
        ? { headerFooterPictureEdits: [...editsByPosition.values()] }
        : {}),
    })
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appHfUpdated'))
  })
  return null
}

/// What the print pane asks of the payload besides the sheet's own Page Layout.
export interface PrintScopeOptions {
  /// Which sheets print: the active one, every visible sheet, or just the
  /// selected range of the active one.
  readonly scope: 'active' | 'workbook' | 'selection'
  /// Excel's "Ignore Print Area": print the whole used range instead.
  readonly ignorePrintArea?: boolean
  /// First / last page of the active sheet's print (1-based).
  readonly pageFrom?: number
  readonly pageTo?: number
  /// Print-only options (custom margins, centering, errors as, B/W, draft, first page number).
  readonly options?: PrintOnlyOptions
}

/// The effective Page Layout of one sheet (file settings + this session's edits).
export function effectiveSetupFor(
  ctx: PageLayoutContext,
  sheetId: string,
): ReturnType<typeof resolveEffectivePageSetup> {
  const state = ctx.lazyWorkbookRef.current
  const journal = state?.editJournal.pageSetup.get(sheetId) ?? {}
  const fileSetup = state?.sheetFilePageSetups.get(sheetId) ?? null
  const fileSheet = state?.file.sheets.find((sheet) => sheet.id === sheetId)
  return resolveEffectivePageSetup(
    journal,
    fileSetup,
    {
      ...(fileSheet?.printArea === undefined ? {} : { printArea: fileSheet.printArea }),
      ...(fileSheet?.printTitles === undefined ? {} : { printTitles: fileSheet.printTitles }),
    },
    state?.editJournal.structuralOps.get(sheetId) ?? [],
  )
}

/// The A1 range of the active selection, or null.
function activeSelectionArea(runtime: UniverRuntime): string | null {
  const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range) return null
  const startColumn = range.getColumn()
  const startRow = range.getRow()
  return (
    `${columnLabel(startColumn)}${startRow + 1}` +
    `:${columnLabel(startColumn + range.getWidth() - 1)}${startRow + range.getHeight()}`
  )
}

/// One sheet laid out as print HTML with its Page Layout settings.
async function sheetPrintPayload(
  ctx: PageLayoutContext,
  worksheet: { getSheetId(): string; getSheetName(): string },
  override?: { readonly printAreas?: readonly string[]; readonly options?: PrintOnlyOptions },
): Promise<WorkbookExportPdfRequest> {
  const state = ctx.lazyWorkbookRef.current
  const sheetId = worksheet.getSheetId()
  const baseSetup = effectiveSetupFor(ctx, sheetId)
  const options = override?.options ?? {}
  const withAreas =
    override?.printAreas === undefined
      ? baseSetup
      : { ...baseSetup, printAreas: override.printAreas }
  const setup =
    options.margins === undefined ? withAreas : { ...withAreas, margins: options.margins }
  const workbookName = state?.file.name ?? 'Book1.xlsx'
  const baseName = workbookName.replace(/\.[^.]+$/, '')
  const pictures = state
    ? await loadHeaderFooterPictures(
        state.file.sessionId,
        setup.headerFooterPictures,
        state.editJournal.pageSetup.get(sheetId)?.headerFooterPictureEdits ?? [],
      )
    : new Map<string, HeaderFooterPictureImage>()
  const frames = await settledVisualFrames(ctx, state, sheetId)
  return buildSheetPrintPayload(
    worksheet as unknown as PrintWorksheet,
    setup,
    `${baseName}.pdf`,
    worksheet.getSheetName(),
    pictures,
    // draft quality prints no drawings
    options.draft === true ? { visuals: [], css: '' } : snapshotPrintVisuals(document, frames),
    workbookName,
    state?.file.csvPath ?? state?.file.path,
    state?.file.themeColors,
    options,
  )
}

/// The active sheet laid out as print HTML with its Page Layout settings, or
/// null (after a status message) when the workbook is not ready for it.
async function activeSheetPrintPayload(
  ctx: PageLayoutContext,
  messages: { readonly notLoaded: string; readonly preparing: string },
): Promise<WorkbookExportPdfRequest | null> {
  const runtime = ctx.univerRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !worksheet) {
    ctx.setMessage(t('appActiveSheetUnavailable'))
    return null
  }
  const state = ctx.lazyWorkbookRef.current
  if (state && !state.flags.preloadComplete) {
    ctx.setMessage(messages.notLoaded)
    return null
  }
  ctx.setMessage(messages.preparing)
  return sheetPrintPayload(ctx, worksheet)
}

/// Lays the active sheet out as HTML with its Page Layout settings and asks
/// the main process to render the PDF (hidden window + save dialog).
/// `outPath` (headless export only) skips the dialog; resolves true when a
/// PDF was written.
export async function handleExportPdf(ctx: PageLayoutContext, outPath?: string): Promise<boolean> {
  try {
    const payload = await activeSheetPrintPayload(ctx, {
      notLoaded: t('appPdfNeedsFullLoad'),
      preparing: t('appPdfRendering'),
    })
    if (!payload) return false
    const result = await window.desktopApi.exportPdf({
      ...payload,
      ...(outPath ? { outPath } : {}),
    })
    ctx.setMessage(
      result.canceled ? t('appPdfCanceled') : t('appPdfExported', { path: result.path }),
    )
    return !result.canceled
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appPdfExportFailed'))
    return false
  }
}

/// The print units of a job — sheets, or row bands of sheets too large to lay out in one piece —
/// each laid out lazily as print HTML (the print sheet's preview and the real print both walk
/// this list). Large workbooks are not loaded whole: a unit reads only its own rows from the file.
/// `error` is a user-facing reason nothing can print.
export type PrintTargets =
  | {
      readonly ok: true
      readonly count: number
      /// Some sheet was split into bands: the real print merges the bands into one PDF.
      readonly banded: boolean
      payloadAt(index: number): Promise<WorkbookExportPdfRequest>
    }
  | { readonly ok: false; readonly error: string }

function a1Of(area: PrintArea): string {
  return `${columnLabel(area.startColumn)}${area.startRow + 1}:${columnLabel(area.endColumn)}${area.endRow + 1}`
}

export function printTargets(ctx: PageLayoutContext, scope: PrintScopeOptions): PrintTargets {
  try {
    const runtime = ctx.univerRef.current
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    const active = workbook?.getActiveSheet()
    if (!runtime || !workbook || !active) {
      return { ok: false, error: t('appActiveSheetUnavailable') }
    }
    const state = ctx.lazyWorkbookRef.current
    const partial = state !== null && !state.flags.preloadComplete
    const sheets =
      scope.scope === 'workbook'
        ? workbook.getSheets().filter((sheet) => !sheet.isSheetHidden())
        : [active]
    let selection: PrintArea | null = null
    if (scope.scope === 'selection') {
      const area = activeSelectionArea(runtime)
      if (area === null) return { ok: false, error: t('appSelectPrintRange') }
      selection = parseArea(area)
    }
    interface Unit {
      readonly sheet: (typeof sheets)[number]
      /// undefined = the sheet's own page layout decides (small, fully loaded sheets)
      readonly areas: readonly PrintArea[] | undefined
    }
    const units: Unit[] = []
    let banded = false
    let cells = 0
    for (const sheet of sheets) {
      const sheetId = sheet.getSheetId()
      const extent = state ? lazySheetScreenExtent(state, sheetId) : null
      const used: PrintArea = {
        startRow: 0,
        startColumn: 0,
        endRow: Math.max(extent ? extent.rows - 1 : sheet.getLastRow(), 0),
        endColumn: Math.max(extent ? extent.columns - 1 : sheet.getLastColumn(), 0),
      }
      const setup = effectiveSetupFor(ctx, sheetId)
      const areas: PrintArea[] =
        selection !== null
          ? [selection]
          : scope.ignorePrintArea === true || setup.printAreas.length === 0
            ? [used]
            : setup.printAreas.map(parseArea)
      cells += totalCells(areas)
      const plan = unitsForAreas(areas)
      if (plan.banded) banded = true
      // explicit areas only when the sheet's own layout would be wrong or too big
      const explicit =
        plan.banded || selection !== null || scope.ignorePrintArea === true || partial
      for (const unit of plan.units) {
        units.push({ sheet, areas: explicit ? unit : undefined })
      }
    }
    if (cells > PRINT_JOB_MAX_CELLS) return { ok: false, error: t('appPrintJobTooLarge') }
    return {
      ok: true,
      count: units.length,
      banded,
      payloadAt: async (index) => {
        const unit = units[index]!
        if (unit.areas !== undefined && partial) {
          for (const area of unit.areas) {
            const resident = await ensureRangeResident(
              runtime,
              ctx.lazyWorkbookRef,
              unit.sheet as unknown as Parameters<typeof ensureRangeResident>[2],
              area,
              ctx.setMessage,
            )
            if (!resident) throw new Error(t('appLoadRangeFailed'))
          }
        }
        return sheetPrintPayload(ctx, unit.sheet, {
          ...(unit.areas === undefined ? {} : { printAreas: unit.areas.map(a1Of) }),
          // a band is a slice of a taller sheet: centering it on its own page would misplace it
          ...(scope.options === undefined
            ? {}
            : {
                options: banded ? { ...scope.options, centerVertically: false } : scope.options,
              }),
        })
      },
    }
  } catch (error: unknown) {
    return { ok: false, error: error instanceof Error ? error.message : t('appPrintFailed') }
  }
}

export interface PrintRunOptions {
  /// Called as units are prepared (band-by-band jobs).
  readonly onProgress?: (done: number, total: number) => void
  readonly signal?: AbortSignal
}

/// File → Print: the same layout, handed to the system print dialog. The
/// print pane's scope decides which sheets / range go out: the active sheet
/// (default), every visible sheet one after the other, or the selection.
export async function handlePrint(
  ctx: PageLayoutContext,
  job?: WorkbookPrintJob,
  scope: PrintScopeOptions = { scope: 'active' },
  run: PrintRunOptions = {},
): Promise<boolean> {
  let bandJob: string | null = null
  try {
    const targets = printTargets(ctx, scope)
    if (!targets.ok) {
      ctx.setMessage(targets.error)
      return false
    }
    ctx.setMessage(t('appPrintPreparing'))
    const pageRange =
      scope.scope !== 'workbook' && scope.pageFrom !== undefined
        ? {
            pageFrom: scope.pageFrom,
            ...(scope.pageTo === undefined ? {} : { pageTo: scope.pageTo }),
          }
        : {}
    if (targets.banded) {
      // very large: every band is laid out, rendered to disk and released in turn, then merged
      const jobId = `print-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
      bandJob = jobId
      for (let index = 0; index < targets.count; index++) {
        if (run.signal?.aborted === true) {
          ctx.setMessage(t('appPrintCanceled'))
          return false
        }
        run.onProgress?.(index, targets.count)
        ctx.setMessage(t('appPrintBandProgress', { done: index + 1, total: targets.count }))
        const added = await window.desktopApi.addPrintBand(jobId, await targets.payloadAt(index))
        if (!added.ok) throw new Error(added.error)
      }
      run.onProgress?.(targets.count, targets.count)
      if (run.signal?.aborted === true) {
        ctx.setMessage(t('appPrintCanceled'))
        return false
      }
      const result = await window.desktopApi.finishPrintBands(jobId, { ...job, ...pageRange })
      bandJob = null
      if (!result.ok) {
        ctx.setMessage(result.error === undefined ? t('appPrintCanceled') : t('appPrintFailed'))
        return false
      }
      ctx.setMessage(result.openedPdf === undefined ? t('appPrintSent') : t('appPrintOpenedPdf'))
      return true
    }
    let sent = false
    for (let index = 0; index < targets.count; index++) {
      const payload = await targets.payloadAt(index)
      const result = await window.desktopApi.printWorkbook(payload, { ...job, ...pageRange })
      if (!result.ok) {
        ctx.setMessage(result.error === undefined ? t('appPrintCanceled') : t('appPrintFailed'))
        return sent
      }
      sent = true
    }
    ctx.setMessage(t('appPrintSent'))
    return sent
  } catch (error: unknown) {
    // layout errors (empty print area, oversized sheet, bad titles) name the cause
    ctx.setMessage(error instanceof Error ? error.message : t('appPrintFailed'))
    return false
  } finally {
    if (bandJob !== null) void window.desktopApi.cancelPrintBands(bandJob)
  }
}

/// The floating visuals of the sheet with their float DOM laid out. Install
/// runs on a timer after load and after viewport changes, so an export that
/// follows the load closely (headless) asks for it and waits for the frames
/// of every visual that is not deleted; visuals without a frame never
/// install, hence the timeout.
async function settledVisualFrames(
  ctx: PageLayoutContext,
  state: LazyWorkbookState | null,
  sheetId: string,
): Promise<readonly InstalledVisualFrame[]> {
  const expected = state
    ? [...state.file.visuals, ...state.editJournal.visualAdds].filter(
        (visual) =>
          visual.sheetId === sheetId && !state.editJournal.visualEdits.get(visual.id)?.remove,
      ).length
    : 0
  if (expected === 0) return []
  if (installedVisualFrames(sheetId).length < expected) {
    ctx.requestVisualInstall?.()
    const deadline = Date.now() + 3000
    while (Date.now() < deadline && installedVisualFrames(sheetId).length < expected) {
      await new Promise((resolve) => setTimeout(resolve, 60))
    }
  }
  const frames = installedVisualFrames(sheetId)
  await settleVisualNodes(document, frames)
  return frames
}

/// Fetches the file's `&G` header/footer pictures as data URLs, keyed by
/// VML slot. Metafiles rasterize to PNG (Chromium cannot paint EMF/WMF); a
/// picture that fails to load is left out — its `&G` then prints nothing,
/// which is also what Excel shows for a slot without a picture.
async function loadHeaderFooterPictures(
  sessionId: string,
  slots: readonly HeaderFooterPictureSlot[],
  edits: readonly import('./edit-journal').HeaderFooterPictureEdit[],
): Promise<Map<string, HeaderFooterPictureImage>> {
  const pictures = new Map<string, HeaderFooterPictureImage>()
  await Promise.all(
    slots.map(async (slot) => {
      try {
        const pending = slot.id.startsWith('journal:')
          ? edits.find((edit) => edit.position === slot.position)?.image
          : undefined
        const media =
          pending && pending !== null
            ? pending
            : await window.desktopApi.readWorkbookMedia({ sessionId, visualId: slot.id })
        const dataUrl = isMetafileMime(media.mediaType)
          ? await metafileToDataUrl(base64ToBytes(media.base64), media.mediaType)
          : `data:${media.mediaType};base64,${media.base64}`
        if (dataUrl) {
          pictures.set(slot.position, { dataUrl, widthPt: slot.widthPt, heightPt: slot.heightPt })
        }
      } catch (reason: unknown) {
        console.warn(`header/footer picture unavailable (${slot.position})`, reason)
      }
    }),
  )
  return pictures
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}
