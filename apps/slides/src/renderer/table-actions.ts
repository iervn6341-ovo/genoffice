/**
 * Table editing actions extracted from App.tsx: row/column
 * structure, merge/split, resize, cell text commit and Tab navigation.
 * Functions read the latest App state through ActionCtx.
 */
import type { EditParagraph } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'
import type { TableCellSel } from './table-selection'

export async function tableStructureOp(
  ctx: ActionCtx,
  sourceId: string,
  kind: 'insert-row' | 'delete-row' | 'insert-col' | 'delete-col',
  index: number,
  before?: boolean,
  /** how many rows / columns at once (the selected range spans that many) */
  count?: number,
  /** cell selection to restore once the table has its new id (default: none) */
  next?: (newSourceId: string) => TableCellSel | null,
): Promise<void> {
  const r = await window.slidesApi.tableStructure({
    slideIndex: ctx.current,
    sourceId,
    kind,
    index,
    ...(before ? { before: true } : {}),
    ...(count && count > 1 ? { count } : {}),
  })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setEditingCell(null)
    ctx.setCellSel(next ? next(r.sourceId) : null)
  } else {
    ctx.setStatus(
      kind.startsWith('delete')
        ? t('appStatusTableDeleteLimit')
        : t('appStatusTableMergedUnsupported'),
    )
  }
}

export async function tableMergeOp(
  ctx: ActionCtx,
  sourceId: string,
  kind: 'merge-right' | 'merge-down' | 'split' | 'merge-range' | 'split-range',
  row: number,
  col: number,
  /** merge-range / split-range: the rectangle's size (row/col is its top-left cell) */
  size?: { rows: number; cols: number },
  /** cell selection to restore once the table has its new id (default: none) */
  next?: (newSourceId: string) => TableCellSel | null,
): Promise<void> {
  const r = await window.slidesApi.tableMerge({
    slideIndex: ctx.current,
    sourceId,
    kind,
    row,
    col,
    ...(size ? { rows: size.rows, cols: size.cols } : {}),
  })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setEditingCell(null)
    ctx.setCellSel(next ? next(r.sourceId) : null)
  } else {
    ctx.setStatus(t('appStatusMergeFailed'))
  }
}

/** `cellEdit_<tableId>_<row>_<col>` (the id TextEditOverlay's cell pseudo-node carries) → the cell */
export function parseCellEditKey(
  key: string | undefined,
): { sourceId: string; row: number; col: number } | null {
  const m = /^cellEdit_(.+)_(\d+)_(\d+)$/.exec(key ?? '')
  return m ? { sourceId: m[1]!, row: Number(m[2]), col: Number(m[3]) } : null
}

export const cellEditKey = (c: { sourceId: string; row: number; col: number }): string =>
  `cellEdit_${c.sourceId}_${c.row}_${c.col}`

/**
 * Write the cell editor's text back. `owner` names the cell the editor DOM was populated for
 * (the overlay is reused when the target cell changes, so a late blur / press-guard commit may
 * run after the current editing cell is already another one — the text still belongs to its
 * own cell, and the newer edit session must survive).
 */
export async function commitCellEdit(
  ctx: ActionCtx,
  paragraphs: EditParagraph[],
  owner?: string,
): Promise<void> {
  const target = parseCellEditKey(owner) ?? ctx.editingCell
  if (!target) return
  const updated = await window.slidesApi.editTableCell({
    slideIndex: ctx.current,
    sourceId: target.sourceId,
    row: target.row,
    col: target.col,
    paragraphs,
  })
  if (updated) ctx.applySlide(ctx.current, updated)
  endCellEdit(ctx, target)
}

/** Leave cell editing, unless the editor has meanwhile moved on to another cell */
export function endCellEdit(
  ctx: Pick<ActionCtx, 'setEditingCell'>,
  cell: { sourceId: string; row: number; col: number },
): void {
  ctx.setEditingCell((cur) =>
    cur && (cur.sourceId !== cell.sourceId || cur.row !== cell.row || cur.col !== cell.col)
      ? cur
      : null,
  )
}

/** Cell Tab/Shift+Tab: commit the current cell (if content changed) and jump to the next/previous cell;
 * Tab on the last cell auto-adds a row and enters its first cell. */
export async function navigateCell(
  ctx: ActionCtx,
  paragraphs: EditParagraph[] | null,
  dir: 1 | -1,
): Promise<void> {
  const { editingCell, slide, current } = ctx
  if (!editingCell || !slide) return
  const tbl = slide.nodes.find((n) => n.sourceId === editingCell.sourceId)
  if (!tbl || tbl.type !== 'table') return
  if (paragraphs) {
    const updated = await window.slidesApi.editTableCell({
      slideIndex: current,
      sourceId: editingCell.sourceId,
      row: editingCell.row,
      col: editingCell.col,
      paragraphs,
    })
    if (updated) ctx.applySlide(current, updated)
  }
  const cells = [...tbl.cells].sort((a, b) => a.row - b.row || a.col - b.col)
  const idx = cells.findIndex((c) => c.row === editingCell.row && c.col === editingCell.col)
  const next = cells[idx + dir]
  if (next) {
    ctx.setEditingCell({ sourceId: editingCell.sourceId, row: next.row, col: next.col })
    return
  }
  if (dir > 0) {
    // Tab on the last cell: add a row and enter its first cell
    const lastRow = Math.max(...cells.map((c) => c.row))
    const r = await window.slidesApi.tableStructure({
      slideIndex: current,
      sourceId: editingCell.sourceId,
      kind: 'insert-row',
      index: lastRow,
    })
    if (r) {
      ctx.applySlide(current, r.slide)
      ctx.setSelectedIds([r.sourceId])
      const firstCol = Math.min(...cells.filter((c) => c.row === lastRow).map((c) => c.col))
      ctx.setEditingCell({ sourceId: r.sourceId, row: lastRow + 1, col: firstCol })
    }
  }
  // Shift+Tab on the first cell: stay in place
}
