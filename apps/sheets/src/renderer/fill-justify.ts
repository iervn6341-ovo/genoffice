/**
 * Home → Fill → Justify. Excel's Range.Justify (read through COM) takes the text of a column of cells,
 * joins it with spaces, and breaks it again into lines no wider than the selection — the sum of its
 * column widths — in the selection's first column, from its first cell down. Cells under the
 * selection take the lines that do not fit (Excel asks first; here the room must be empty), cells of
 * the selection left over are emptied, and a selection with a number or a formula in it is refused.
 */
import type { ICellData, IRange } from '@univerjs/core'

import { t } from './i18n/locale'
import { FILL_CELL_CAP } from './fill-across'
import { ensureRangeResident } from './lazy-resident'
import type { RibbonCommandContext } from './ribbon-actions'
import { runAsOneUndoStepAsync } from './undo-group'
import type { UniverRuntime } from './univer-state'

/** the lines `text` breaks into at spaces, none wider than `width` unless one word alone is */
export function wrapWords(
  text: string,
  width: number,
  measure: (text: string) => number,
): string[] {
  const words = text.split(/\s+/).filter((word) => word !== '')
  const lines: string[] = []
  let line = ''
  for (const word of words) {
    const next = line === '' ? word : `${line} ${word}`
    if (line === '' || measure(next) <= width) {
      line = next
    } else {
      lines.push(line)
      line = word
    }
  }
  if (line !== '') lines.push(line)
  return lines
}

export type JustifyError = 'noText' | 'notText' | 'below' | 'rich'

export interface JustifyPlan {
  /** the lines, written into the first column from the first row */
  readonly lines: readonly string[]
  /** rows of the selection the lines do not reach: emptied */
  readonly cleared: number
  /** rows the lines need under the selection */
  readonly extra: number
}

/**
 * The lines of a selection's text and what they do to the rows. `cells` is the selection row by row
 * (each row: its cells), `measure` the width of a text in the cell's font, `width` the selection's.
 */
export function planJustify(
  cells: readonly (readonly { v?: unknown; f?: unknown; p?: unknown }[])[],
  width: number,
  measure: (text: string) => number,
): JustifyPlan | { readonly error: JustifyError } {
  const first: string[] = []
  for (const row of cells) {
    for (const [column, cell] of row.entries()) {
      const hasFormula = typeof cell.f === 'string' && cell.f !== ''
      // formatted text (rich text, a link with its own runs, line breaks) is not in `v`: joining the
      // plain text would drop it and the rewrite would erase the cell
      if (cell.p !== undefined && cell.p !== null) return { error: 'rich' }
      if (hasFormula || typeof cell.v === 'number' || typeof cell.v === 'boolean') {
        return { error: 'notText' }
      }
      if (column === 0 && typeof cell.v === 'string' && cell.v.trim() !== '') first.push(cell.v)
    }
  }
  if (first.length === 0) return { error: 'noText' }
  const lines = wrapWords(first.join(' '), width, measure)
  return {
    lines,
    cleared: Math.max(0, cells.length - lines.length),
    extra: Math.max(0, lines.length - cells.length),
  }
}

/** horizontal space a cell keeps free of text, like Excel's cell padding */
const CELL_PADDING = 6

function textMeasure(style: {
  fs?: number
  ff?: string
  bl?: number
  it?: number
}): (text: string) => number {
  const canvas = document.createElement('canvas').getContext('2d')
  const size = Math.round(((style.fs ?? 11) * 96) / 72)
  const font = `${style.it === 1 ? 'italic ' : ''}${style.bl === 1 ? 'bold ' : ''}${size}px ${
    style.ff ?? 'Calibri'
  }, sans-serif`
  if (canvas) canvas.font = font
  return (text) => (canvas ? canvas.measureText(text).width : text.length * size * 0.55)
}

export async function justifySelection(
  ctx: RibbonCommandContext,
  runtime: UniverRuntime,
): Promise<void> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !sheet || !active) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const range: IRange = {
    startRow: active.getRow(),
    endRow: active.getRow() + active.getHeight() - 1,
    startColumn: active.getColumn(),
    endColumn: active.getColumn() + active.getWidth() - 1,
  }
  if (active.getHeight() * active.getWidth() > FILL_CELL_CAP) {
    ctx.setMessage(t('appCommandTooBig'))
    return
  }
  const resident = await ensureRangeResident(
    runtime,
    ctx.lazyWorkbookRef,
    sheet,
    range,
    ctx.setMessage,
  )
  if (!resident) {
    ctx.setMessage(t('appAutofillStreaming'))
    return
  }
  const cells = sheet
    .getRange(range.startRow, range.startColumn, active.getHeight(), active.getWidth())
    .getCellDatas()
  let width = 0
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    width += sheet.getColumnWidth(column)
  }
  const style = sheet.getRange(range.startRow, range.startColumn, 1, 1).getCellStyleData() ?? {}
  const plan = planJustify(
    cells.map((row) => row.map((cell) => cell ?? {})),
    Math.max(1, width - CELL_PADDING),
    textMeasure(style as { fs?: number; ff?: string; bl?: number; it?: number }),
  )
  if ('error' in plan) {
    ctx.setMessage(
      t(
        plan.error === 'notText'
          ? 'appJustifyNotText'
          : plan.error === 'rich'
            ? 'appJustifyRich'
            : plan.error === 'noText'
              ? 'appJustifyNeedsText'
              : 'appJustifyBelow',
      ),
    )
    return
  }
  // the lines that do not fit the selection go under it, where there must be room
  if (plan.extra > 0) {
    const below: IRange = {
      startRow: range.endRow + 1,
      endRow: range.endRow + plan.extra,
      startColumn: range.startColumn,
      endColumn: range.startColumn,
    }
    if (below.endRow >= sheet.getMaxRows()) {
      ctx.setMessage(t('appJustifyBelow'))
      return
    }
    if (!(await ensureRangeResident(runtime, ctx.lazyWorkbookRef, sheet, below, ctx.setMessage))) {
      ctx.setMessage(t('appAutofillStreaming'))
      return
    }
    const there = sheet
      .getRange(below.startRow, below.startColumn, plan.extra, 1)
      .getCellDatas()
      .flat()
    if (there.some((cell) => cell && ((cell.v ?? '') !== '' || cell.f || cell.p))) {
      ctx.setMessage(t('appJustifyBelow'))
      return
    }
  }
  const matrix: Record<number, Record<number, ICellData>> = {}
  plan.lines.forEach((line, index) => {
    ;(matrix[range.startRow + index] ??= {})[range.startColumn] = {
      v: line,
      t: 1,
      f: null,
      si: null,
      p: null,
    }
  })
  for (let row = range.startRow + plan.lines.length; row <= range.endRow; row += 1) {
    ;(matrix[row] ??= {})[range.startColumn] = { v: null, f: null, si: null, p: null }
  }
  await runAsOneUndoStepAsync(runtime, async () => {
    await runtime.univerAPI.executeCommand('sheet.command.set-range-values', {
      unitId: workbook.getId(),
      subUnitId: sheet.getSheetId(),
      value: matrix,
    })
  })
  ctx.setMessage(t('appFillJustified'))
}
