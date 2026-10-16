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

import { t } from './i18n/locale'
import { ensureRangeResident } from './lazy-resident'
import { mergeConditionalFormats, pasteMerges, pasteNotes, pasteValidation } from './paste-extras'
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
import { readClipboardCopy } from './paste-source'
import { sharedFormulaResolverFor } from './shared-formula-journal'
import type { RibbonCommandContext } from './ribbon-actions'
import { runAsOneUndoStepAsync } from './undo-group'
import type { UniverRuntime } from './univer-state'

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
  cell.p = null
  if (write.value.v !== null) {
    cell.t = write.value.t ?? (typeof write.value.v === 'string' ? CellValueType.STRING : undefined)
  }
  return cell
}

export async function pasteSpecial(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  options: PasteSpecialOptions,
): Promise<void> {
  if (!needsOwnPaste(options)) {
    const command = NATIVE_PASTE_COMMANDS[options.what]
    if (command) {
      // an explicit params object is needed by the plain paste (see the Paste button)
      void runtime.univerAPI.executeCommand(
        command,
        command === 'univer.command.paste' ? {} : undefined,
      )
    }
    return
  }
  // what the system clipboard holds, as Univer's own paste decides it (not just the last sheet copy)
  const copied = await readClipboardCopy(runtime)
  if (!copied) {
    ctx.setMessage(t('appPasteSpecialNeedsCopy'))
    return
  }
  if (copied.cut) {
    ctx.setMessage(t('appPasteSpecialNotAfterCut'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !sheet || !active) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
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
  if (
    extent.endRow >= sheet.getMaxRows() ||
    extent.endColumn >= sheet.getMaxColumns() ||
    cellCount > MAX_PASTE_CELLS
  ) {
    ctx.setMessage(t('appPasteSpecialTooLarge'))
    return
  }
  const resident = await ensureRangeResident(
    runtime,
    ctx.lazyWorkbookRef,
    sheet,
    extent,
    ctx.setMessage,
  )
  if (!resident) {
    ctx.setMessage(t('appAutofillStreaming'))
    return
  }

  const lexer = runtime.univer.__getInjector().get(LexerTreeBuilder)
  const sourceFormulas = sharedFormulaResolverFor(runtime, copied.sheetId)
  const destinationFormulas = sharedFormulaResolverFor(runtime, sheet.getSheetId())
  const datas = sheet
    .getRange(
      extent.startRow,
      extent.startColumn,
      extent.endRow - extent.startRow + 1,
      extent.endColumn - extent.startColumn + 1,
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
          const f =
            cell.f ||
            (cell.si ? sourceFormulas(copied.rows[i]!, copied.columns[j]!, cell.si) : null)
          // a rich-text cell has no plain value: paste its text
          const plain =
            (cell.v === undefined || cell.v === null || cell.v === '') && cell.p
              ? extractPureTextFromCell(cell as CellData)
              : ''
          return {
            v: plain ? plain : (cell.v ?? null),
            t: plain ? CellValueType.STRING : (cell.t ?? null),
            f: f ?? null,
            style: typeof cell.s === 'object' ? cell.s : null,
          }
        }),
      ),
    },
    options,
    target,
    sameSheet: copied.sheetId === sheet.getSheetId(),
    ...(() => {
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
    shiftFormula: (formula, columns, rows) => lexer.moveFormulaRefOffset(formula, columns, rows),
  })
  if ('error' in plan) {
    ctx.setMessage(
      t(plan.error === 'overlap' ? 'appPasteSpecialOverlap' : 'appPasteSpecialNeedsCopy'),
    )
    return
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
  const sourceSheet = workbook.getSheetBySheetId(copied.sheetId)
  const extras = sourceSheet ? pasteExtras(options.what) : null
  await runAsOneUndoStepAsync(runtime, async () => {
    const run = (value: unknown) =>
      runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
        unitId,
        subUnitId,
        value,
      })
    if (Object.keys(clear).length > 0) await run(clear)
    if (Object.keys(apply).length > 0) await run(apply)
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
      if (ready(sheet.getSheetId()) && ready(sourceSheet.getSheetId())) {
        await pasteValidation(runtime, sheet, sourceSheet, pastedCells(copied, options, target))
      } else {
        ctx.setMessage(t('appSheetStillIndexing'))
      }
    }
    if (options.what === 'allMergingConditionalFormats') {
      mergeConditionalFormats(sheet, sourceSheet, copied, options, target)
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
}
