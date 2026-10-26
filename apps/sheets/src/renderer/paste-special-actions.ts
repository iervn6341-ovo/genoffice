/**
 * Home → Paste ▾ / Paste Special…: runs Univer's own paste variants where they match Excel, and
 * the plan in paste-special.ts for Transpose, Operation, Skip blanks and the number-format
 * variants (see there for the behavior). The plan's writes go through `set-range-values`, so the
 * edit journals, saves and undoes like any cell edit — one ⌘Z for the whole paste.
 */
import { ensureColumnOnActiveWorkbook } from './column-extension'
import type { ICellData, IStyleData } from '@univerjs/core'
import { CellValueType, extractPureTextFromCell, type ICellData as CellData } from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import { ISheetClipboardService } from '@univerjs/sheets-ui'

import { t } from './i18n/locale'
import { ensureRangeResident } from './lazy-resident'
import {
  mergeConditionalFormats,
  pasteColumnWidths,
  pasteLinks,
  pasteMerges,
  pasteNotes,
  pasteValidation,
} from './paste-extras'
import {
  NATIVE_PASTE_COMMANDS,
  needsOwnPaste,
  pasteExtent,
  pasteExtras,
  pastedCells,
  planPaste,
  type PasteSpecialOptions,
  type PlannedWrite,
} from './paste-special'
import { foreignFormula, virtualSourceSheet, type ForeignCopy } from './clipboard-foreign'
import { rowBatches } from './paste-batches'
import { readClipboardSource, type CopiedCells } from './paste-source'
import { sharedFormulaResolverFor } from './shared-formula-journal'
import type { RibbonCommandContext } from './ribbon-actions'
import { runAsOneUndoStepAsync } from './undo-group'
import type { UniverRuntime, UniverWorksheet } from './univer-state'

/// Excel's last row (1-based 1,048,576)
const MAX_SHEET_ROWS = 1_048_576

/// a paste bigger than this is refused: the plan reads the destination cell by cell
const MAX_PASTE_CELLS = 250_000

const PASTE_WHATS = new Set([
  'all',
  'formulas',
  'values',
  'formats',
  'allExceptBorders',
  'columnWidths',
  'formulasAndNumberFormats',
  'valuesAndNumberFormats',
  'comments',
  'validation',
  'allUsingSourceTheme',
  'allMergingConditionalFormats',
])
const PASTE_OPERATIONS = new Set(['none', 'add', 'subtract', 'multiply', 'divide'])

/// The dialog's choice as it travels in the command string; null when it does not read as one
export function parsePasteSpecialOptions(encoded: string): PasteSpecialOptions | null {
  try {
    const raw = JSON.parse(decodeURIComponent(encoded)) as Partial<PasteSpecialOptions>
    if (
      typeof raw.what !== 'string' ||
      !PASTE_WHATS.has(raw.what) ||
      typeof raw.operation !== 'string' ||
      !PASTE_OPERATIONS.has(raw.operation) ||
      typeof raw.skipBlanks !== 'boolean' ||
      typeof raw.transpose !== 'boolean'
    ) {
      return null
    }
    return raw as PasteSpecialOptions
  } catch {
    return null
  }
}

/// the value side of a planned write as the set-range-values cell
function valueCell(write: PlannedWrite): Record<string, unknown> {
  const cell: Record<string, unknown> = {}
  if (!write.value) return cell
  cell.v = write.value.v
  cell.f = write.value.f
  cell.si = null
  cell.p = write.value.p ?? null
  if (write.value.v !== null) {
    cell.t = write.value.t ?? (typeof write.value.v === 'string' ? CellValueType.STRING : undefined)
  }
  return cell
}

/// the choices that bring the source's look (and with it conditional formats and links)
const carriesFormats = (what: PasteSpecialOptions['what']): boolean =>
  what === 'all' ||
  what === 'allExceptBorders' ||
  what === 'allUsingSourceTheme' ||
  what === 'allMergingConditionalFormats'

/// Univer's paste hooks may refuse a paste (protected range, too many cells) and say why
function pasteAllowed(
  runtime: UniverRuntime,
  unitId: string,
  subUnitId: string,
  extent: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  try {
    const hooks = runtime.univer.__getInjector().get(ISheetClipboardService).getClipboardHooks()
    const range = {
      rows: Array.from(
        { length: extent.endRow - extent.startRow + 1 },
        (_, i) => extent.startRow + i,
      ),
      cols: Array.from(
        { length: extent.endColumn - extent.startColumn + 1 },
        (_, i) => extent.startColumn + i,
      ),
    }
    return !hooks.some((hook) => hook.onBeforePaste?.({ unitId, subUnitId, range }) === false)
  } catch {
    return true
  }
}

export async function pasteSpecial(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  options: PasteSpecialOptions,
  /// a copy that came from outside this window (another workbook, Excel, text); read from the
  /// system clipboard when absent
  givenForeign?: ForeignCopy | null,
): Promise<boolean> {
  let foreign: ForeignCopy | null = givenForeign ?? null
  let copied: CopiedCells | null = foreign?.copied ?? null
  if (!foreign) {
    // what the system clipboard holds, as Univer's own paste decides it (not just the last sheet copy)
    const source = await readClipboardSource(runtime).catch(() => null)
    foreign = source?.foreign ?? null
    copied = source?.copied ?? null
    // a block copied in another window is not in Univer's cache: only the plan below can paste it
    if (!foreign && !needsOwnPaste(options)) {
      const command = NATIVE_PASTE_COMMANDS[options.what]
      if (command) {
        // an explicit params object is needed by the plain paste (see the Paste button)
        void runtime.univerAPI.executeCommand(
          command,
          command === 'univer.command.paste' ? {} : undefined,
        )
      }
      return false
    }
  }
  if (!copied) {
    ctx.setMessage(t('appPasteSpecialNeedsCopy'))
    return false
  }
  // after Cut only the plain paste works (Range.PasteSpecial throws in Excel)
  const plainPaste =
    options.what === 'all' &&
    options.operation === 'none' &&
    !options.skipBlanks &&
    !options.transpose
  if (copied.cut || (foreign?.cutCopyId && !plainPaste)) {
    ctx.setMessage(t('appPasteSpecialNotAfterCut'))
    return false
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !sheet || !active) {
    ctx.setMessage(t('appSelectCellFirst'))
    return false
  }
  const target = {
    anchorRow: active.getRow(),
    anchorColumn: active.getColumn(),
    selectionRows: active.getHeight(),
    selectionColumns: active.getWidth(),
  }
  const extent = pasteExtent(copied.rows.length, copied.columns.length, options, target)
  const cellCount =
    (extent.endRow - extent.startRow + 1) * (extent.endColumn - extent.startColumn + 1)
  // Columns run to XFD: grow the grid first so a paste near the edge fits.
  if (cellCount <= MAX_PASTE_CELLS) {
    ensureColumnOnActiveWorkbook(runtime, sheet.getSheetId(), extent.endColumn)
  }
  // the sheet's own gatekeepers (permissions, size limit) get their say, as in Univer's paste
  if (!pasteAllowed(runtime, workbook.getId(), sheet.getSheetId(), extent)) return false
  // Rows run to 1,048,576 like Excel's; a sheet with fewer rows grows to hold the paste (inside
  // the paste's undo step, below), as Univer's own paste does.
  const sheetRows = sheet.getMaxRows()
  if (
    extent.endRow >= MAX_SHEET_ROWS ||
    extent.endColumn >= sheet.getMaxColumns() ||
    cellCount > MAX_PASTE_CELLS
  ) {
    ctx.setMessage(t('appPasteSpecialTooLarge'))
    return false
  }
  // rows past the sheet's end hold nothing to read yet
  const existing = { ...extent, endRow: Math.min(extent.endRow, sheetRows - 1) }
  const resident =
    existing.endRow < existing.startRow ||
    (await ensureRangeResident(runtime, ctx.lazyWorkbookRef, sheet, existing, ctx.setMessage))
  if (!resident) {
    ctx.setMessage(t('appAutofillStreaming'))
    return false
  }

  const lexer = runtime.univer.__getInjector().get(LexerTreeBuilder)
  const knownNames = new Set<string>()
  try {
    for (const name of workbook.getDefinedNames()) knownNames.add(name.getName().toUpperCase())
  } catch {
    // names that cannot be listed count as unknown
  }
  const sourceFormulas = foreign ? () => null : sharedFormulaResolverFor(runtime, copied.sheetId)
  const destinationFormulas = sharedFormulaResolverFor(runtime, sheet.getSheetId())
  const datas =
    existing.endRow < existing.startRow
      ? []
      : sheet
          .getRange(
            existing.startRow,
            existing.startColumn,
            existing.endRow - existing.startRow + 1,
            existing.endColumn - existing.startColumn + 1,
          )
          .getCellDatas()
  const plan = planPaste({
    source: {
      rows: copied.rows,
      columns: copied.columns,
      cells: copied.cells.map((line, i) =>
        line.map((cell, j) => {
          if (cell === null) return null
          // a filled / tiled cell is a shared-formula follower with no formula text of its own
          const own =
            cell.f ||
            (cell.si ? sourceFormulas(copied.rows[i]!, copied.columns[j]!, cell.si) : null)
          // a formula that points at another sheet of its own workbook cannot follow the block
          // into another workbook: it lands as its last result
          const f = foreign && !foreign.literalFormulas ? foreignFormula(own, knownNames) : own
          // a rich-text cell has no plain value: paste its text
          const plain =
            (cell.v === undefined || cell.v === null || cell.v === '') && cell.p
              ? extractPureTextFromCell(cell as CellData)
              : ''
          return {
            v: plain ? plain : (cell.v ?? null),
            t: plain ? CellValueType.STRING : (cell.t ?? null),
            f: f ?? null,
            ...(cell.p ? { p: cell.p } : {}),
            style: typeof cell.s === 'object' ? cell.s : null,
          }
        }),
      ),
    },
    options,
    target,
    sameSheet: foreign ? false : copied.sheetId === sheet.getSheetId(),
    ...(() => {
      if (foreign) return foreign.sheetName ? { sourceSheetName: foreign.sheetName } : {}
      const name = workbook.getSheetBySheetId(copied.sheetId)?.getSheetName()
      return name === undefined ? {} : { sourceSheetName: name }
    })(),
    readDestination: (row, column) => {
      const cell = datas[row - extent.startRow]?.[column - extent.startColumn]
      // a rich-text cell has no plain value, but it is text, not a blank the operation could fill with 0
      const v =
        (cell?.v === undefined || cell.v === null || cell.v === '') && cell?.p
          ? extractPureTextFromCell(cell as CellData)
          : cell?.v
      return {
        v: typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' ? v : null,
        f: cell?.f || (cell?.si ? destinationFormulas(row, column, cell.si) : null) || null,
      }
    },
    readDestinationStyle: (row, column) => sheet.getRange(row, column, 1, 1).getCellStyleData(),
    shiftFormula: (formula, columns, rows) =>
      foreign?.literalFormulas ? formula : lexer.moveFormulaRefOffset(formula, columns, rows),
  })
  if ('error' in plan) {
    ctx.setMessage(
      t(plan.error === 'overlap' ? 'appPasteSpecialOverlap' : 'appPasteSpecialNeedsCopy'),
    )
    return false
  }

  // A format replaces the old one: `{ s: null }` clears it first, because Univer merges a style
  // object key by key and would leave the destination's own settings standing.
  const clear: Record<number, Record<number, ICellData>> = {}
  const apply: Record<number, Record<number, ICellData>> = {}
  for (const write of plan.writes) {
    const cell = valueCell(write)
    if (write.style?.mode === 'replace') {
      ;(clear[write.row] ??= {})[write.column] = { s: null } as unknown as ICellData
      if (write.style.style) cell.s = write.style.style as IStyleData
    } else if (write.style?.mode === 'numfmt') {
      // a null pattern takes the number format off (back to General)
      cell.s = { n: write.style.pattern === null ? null : { pattern: write.style.pattern } }
    }
    if (Object.keys(cell).length > 0) (apply[write.row] ??= {})[write.column] = cell as ICellData
  }
  const unitId = workbook.getId()
  const subUnitId = sheet.getSheetId()
  // what the choice carries beside values and formats comes from the sheet the block was copied
  // from; when that sheet is gone only the cells are pasted
  const sourceSheet = foreign
    ? (virtualSourceSheet(foreign) as unknown as UniverWorksheet)
    : workbook.getSheetBySheetId(copied.sheetId)
  const extras = sourceSheet ? pasteExtras(options.what) : null
  await runAsOneUndoStepAsync(runtime, async () => {
    const run = (value: unknown) =>
      runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
        unitId,
        subUnitId,
        value,
      })
    if (extent.endRow >= sheetRows) {
      await sheet.insertRowsAfter(sheetRows - 1, extent.endRow - sheetRows + 1)
    }
    for (const part of rowBatches(clear)) await run(part)
    for (const part of rowBatches(apply)) await run(part)
    if (!sourceSheet || !extras) return
    if (extras.merges) {
      await pasteMerges(sheet, sourceSheet, copied, options, target, plan.extent)
    }
    if (extras.notes)
      await pasteNotes(ctx, sheet, sourceSheet, pastedCells(copied, options, target))
    if (extras.validation) {
      // a sheet whose file validations are not installed yet reads as having none: pasting then
      // would clear the destination's rules and add nothing
      const state = ctx.lazyWorkbookRef.current
      const ready = (id: string): boolean =>
        !state || state.editJournal.sheets.added.has(id) || state.appliedDvSheets.has(id)
      if (ready(sheet.getSheetId()) && (foreign || ready(sourceSheet.getSheetId()))) {
        await pasteValidation(runtime, sheet, sourceSheet, pastedCells(copied, options, target))
      } else {
        ctx.setMessage(t('appSheetStillIndexing'))
      }
    }
    if (
      options.what === 'allMergingConditionalFormats' ||
      (foreign && carriesFormats(options.what))
    ) {
      mergeConditionalFormats(sheet, sourceSheet, copied, options, target)
    }
    if (foreign && foreign.links.length > 0 && carriesFormats(options.what)) {
      await pasteLinks(ctx, runtime, sheet, foreign.links, pastedCells(copied, options, target))
    }
    if (foreign && options.what === 'columnWidths') {
      pasteColumnWidths(sheet, foreign, copied, options, target)
    }
  })
  sheet
    .getRange(
      plan.extent.startRow,
      plan.extent.startColumn,
      plan.extent.endRow - plan.extent.startRow + 1,
      plan.extent.endColumn - plan.extent.startColumn + 1,
    )
    .activate()
  return true
}
