/**
 * Home → Fill → Across Worksheets. Excel does it for the sheets that are grouped (the tabs
 * selected together); the tabs here are not grouped, so the dialog picks the sheets. The cells of the
 * selection are copied to the same cells of each picked sheet — All (contents and formats), Contents
 * or Formats — with formulas copied as they are (their relative references stay relative, so a
 * formula on the other sheet reads that sheet's cells), in one undo step.
 */
import {
  IUniverInstanceService,
  UniverInstanceType,
  type ICellData,
  type IRange,
  type IStyleData,
  type Workbook,
} from '@univerjs/core'

import { t } from './i18n/locale'
import { ensureRangeResident } from './lazy-resident'
import type { RibbonCommandContext } from './ribbon-actions'
import { sharedFormulaResolverFor } from './shared-formula-journal'
import { runAsOneUndoStepAsync } from './undo-group'
import type { UniverRuntime } from './univer-state'

/// the most cells Across Worksheets and Justify read and rewrite in one command
export const FILL_CELL_CAP = 250_000

export type FillAcrossWhat = 'all' | 'contents' | 'formats'

export interface FillAcrossParams {
  readonly sheetIds: readonly string[]
  readonly what: FillAcrossWhat
}

/** the sheets the selection can be filled to: every other one that is shown */
export function otherSheets(runtime: UniverRuntime): { id: string; name: string }[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const active = workbook?.getActiveSheet()?.getSheetId()
  return (workbook?.getSheets() ?? [])
    .filter((sheet) => sheet.getSheetId() !== active && !sheet.isSheetHidden())
    .map((sheet) => ({ id: sheet.getSheetId(), name: sheet.getSheetName() }))
}

export async function fillAcrossSheets(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
  params: FillAcrossParams,
): Promise<void> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const source = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !source || !active) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const range: IRange = {
    startRow: active.getRow(),
    endRow: active.getRow() + active.getHeight() - 1,
    startColumn: active.getColumn(),
    endColumn: active.getColumn() + active.getWidth() - 1,
  }
  // a whole column or row is ~1M cells per sheet: refuse before reading them
  if (active.getHeight() * active.getWidth() > FILL_CELL_CAP) {
    ctx.setMessage(t('appCommandTooBig'))
    return
  }
  const targets = params.sheetIds
    .map((id) => workbook.getSheetBySheetId(id))
    .filter((sheet): sheet is NonNullable<typeof sheet> => sheet !== null && sheet !== undefined)
    .filter((sheet) => sheet.getSheetId() !== source.getSheetId())
  if (targets.length === 0) {
    ctx.setMessage(t('appFillAcrossNoSheets'))
    return
  }
  const resident = await ensureRangeResident(
    runtime,
    ctx.lazyWorkbookRef,
    source,
    range,
    ctx.setMessage,
  )
  if (!resident) {
    ctx.setMessage(t('appAutofillStreaming'))
    return
  }
  for (const target of targets) {
    await ensureRangeResident(runtime, ctx.lazyWorkbookRef, target, range, ctx.setMessage)
  }

  const styles = runtime.univer
    .__getInjector()
    .get(IUniverInstanceService)
    .getUnit<Workbook>(workbook.getId(), UniverInstanceType.UNIVER_SHEET)
    ?.getStyles()
  const formulas = sharedFormulaResolverFor(runtime, source.getSheetId())
  const cells = source
    .getRange(range.startRow, range.startColumn, active.getHeight(), active.getWidth())
    .getCellDatas()
  const withContents = params.what !== 'formats'
  const withFormats = params.what !== 'contents'

  const clear: Record<number, Record<number, ICellData>> = {}
  const apply: Record<number, Record<number, ICellData>> = {}
  cells.forEach((line, i) => {
    line.forEach((cell, j) => {
      const row = range.startRow + i
      const column = range.startColumn + j
      const next: ICellData = {}
      if (withContents) {
        const formula = cell?.f || (cell?.si ? formulas(row, column, cell.si) : null) || null
        next.v = cell?.v ?? null
        next.f = formula
        next.si = null
        next.p = cell?.p ?? null
        if (cell?.t !== undefined && cell.t !== null) next.t = cell.t
      }
      if (withFormats) {
        ;(clear[row] ??= {})[column] = { s: null } as unknown as ICellData
        const own = typeof cell?.s === 'string' ? styles?.get(cell.s) : cell?.s
        if (own) next.s = own as IStyleData
      }
      if (Object.keys(next).length > 0) (apply[row] ??= {})[column] = next
    })
  })

  await runAsOneUndoStepAsync(runtime, async () => {
    for (const target of targets) {
      const run = (value: unknown) =>
        runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
          unitId: workbook.getId(),
          subUnitId: target.getSheetId(),
          value,
        })
      if (Object.keys(clear).length > 0) await run(clear)
      if (Object.keys(apply).length > 0) await run(apply)
    }
  })
  ctx.setMessage(t('appFillAcrossDone', { count: String(targets.length) }))
}
