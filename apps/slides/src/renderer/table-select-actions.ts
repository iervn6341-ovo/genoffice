/**
 * Table cell-range selection: the commands that act on the selected block (clear, copy / cut /
 * paste, merge / split, insert / delete rows and columns) and the key handling while cells — not
 * text — are selected. State lives in App (`cellSel`); everything here reads the latest through
 * ActionCtx. Selection geometry is in table-selection.ts.
 */
import type { TableRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'
import { runOps } from './table-layout-actions'
import { tableMergeOp, tableStructureOp } from './table-actions'
import {
  allSel,
  cellAddrsInRect,
  cellCovering,
  cellRect,
  extendSel,
  gridCols,
  gridRows,
  moveSel,
  parseTsv,
  rangeToHtml,
  rangeToTsv,
  selectionRect,
  singleCellSel,
  tabTarget,
  type CellRect,
  type CellAddr,
  type Dir,
  type TableCellSel,
} from './table-selection'

/** The table the cell selection lives in (it must also be the one selected element) */
export function selectedTable(ctx: ActionCtx): TableRenderNode | null {
  const sel = ctx.cellSel
  if (!sel || ctx.selectedIds.length !== 1 || ctx.selectedIds[0] !== sel.sourceId) return null
  const n = ctx.slide?.nodes.find((x) => x.sourceId === sel.sourceId)
  return n?.type === 'table' ? (n as TableRenderNode) : null
}

/** Rectangle the selection covers, with its table (null without a cell selection) */
export function currentRange(ctx: ActionCtx): { table: TableRenderNode; rect: CellRect } | null {
  const table = selectedTable(ctx)
  return table && ctx.cellSel ? { table, rect: selectionRect(table, ctx.cellSel) } : null
}

const rectOfCell = (table: TableRenderNode, at: CellAddr): CellRect => {
  const c = cellCovering(table, at.row, at.col)
  return c ? cellRect(c) : { r0: at.row, c0: at.col, r1: at.row, c1: at.col }
}

/**
 * The block a table command should act on: the cell selection, or — from the cell being edited or
 * a right-clicked cell — that one cell.
 */
export function commandRange(
  ctx: ActionCtx,
  table: TableRenderNode,
  at?: CellAddr | null,
): CellRect | null {
  if (ctx.cellSel?.sourceId === table.sourceId) {
    const rect = selectionRect(table, ctx.cellSel)
    // a right-click inside the selection keeps the whole block; outside it targets just that cell
    if (!at || (at.row >= rect.r0 && at.row <= rect.r1 && at.col >= rect.c0 && at.col <= rect.c1)) {
      return rect
    }
  }
  const cell = at ?? (ctx.editingCell?.sourceId === table.sourceId ? ctx.editingCell : null)
  return cell ? rectOfCell(table, cell) : null
}

const target = (ctx: ActionCtx, table: TableRenderNode) => ({
  slide: ctx.current,
  el: table.sourceId,
})

/** Delete / Backspace: empty the selected cells (formatting stays) */
export async function clearSelectedCells(ctx: ActionCtx): Promise<void> {
  const range = currentRange(ctx)
  if (!range) return
  const cells = cellAddrsInRect(range.table, range.rect)
  await runOps(ctx, [{ op: 'clearTableCells', target: target(ctx, range.table), cells }])
}

/** Text of the last copy, for pasting when the system clipboard cannot be read */
let lastCopiedTsv = ''

export async function copySelectedCells(ctx: ActionCtx, cut = false): Promise<void> {
  const range = currentRange(ctx)
  if (!range) return
  const tsv = rangeToTsv(range.table, range.rect)
  lastCopiedTsv = tsv
  try {
    // plain text for editors, an HTML table for spreadsheets / Pages / Numbers (one clipboard write)
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/plain': new Blob([tsv], { type: 'text/plain' }),
        'text/html': new Blob([rangeToHtml(range.table, range.rect)], { type: 'text/html' }),
      }),
    ])
  } catch {
    try {
      await navigator.clipboard.writeText(tsv)
    } catch {
      // the in-memory copy still serves a paste inside this window
    }
  }
  if (cut) await clearSelectedCells(ctx)
}

/**
 * Paste tab/newline separated text (our own copy, or a range copied in Excel) with its top-left
 * field in the first selected cell; the block grows right/down but never past the table. Text
 * only: the destination cells keep their own formatting.
 */
export async function pasteIntoCells(ctx: ActionCtx): Promise<void> {
  const range = currentRange(ctx)
  if (!range) return
  let text: string
  try {
    text = await navigator.clipboard.readText()
  } catch {
    // unreadable system clipboard: fall back to our own last copy. (A readable clipboard with no
    // text — an image, files, or cleared — pastes nothing: reusing an old copy there would drop
    // stale cell text into the selected cells.)
    text = lastCopiedTsv
  }
  if (!text) return
  const rows = parseTsv(text)
  if (!rows.length) return
  const { table, rect } = range
  const ops: unknown[] = []
  for (let i = 0; i < rows.length; i++) {
    for (let j = 0; j < rows[i]!.length; j++) {
      const cell = table.cells.find((c) => c.row === rect.r0 + i && c.col === rect.c0 + j)
      if (!cell) continue
      const paragraphs = rows[i]![j]!.split('\n').map((line) => ({ runs: [{ text: line }] }))
      ops.push({
        op: 'setTableCell',
        target: target(ctx, table),
        row: cell.row,
        col: cell.col,
        paragraphs,
      })
    }
  }
  if (!ops.length) return
  if (ops.length > 50) {
    ctx.setStatus(t('appStatusTablePasteLimit'))
    ops.length = 50
  }
  await runOps(ctx, ops)
  // the pasted block becomes the selection
  const lastRow = Math.min(gridRows(table) - 1, rect.r0 + rows.length - 1)
  const width = Math.max(...rows.map((r) => r.length))
  const lastCol = Math.min(gridCols(table) - 1, rect.c0 + width - 1)
  ctx.setCellSel({
    sourceId: table.sourceId,
    anchor: { row: rect.r0, col: rect.c0 },
    focus: { row: lastRow, col: lastCol },
  })
}

/** Merge Cells: the whole block becomes one cell */
export async function mergeSelectedCells(ctx: ActionCtx): Promise<void> {
  const range = currentRange(ctx)
  if (!range) return
  const { table, rect } = range
  const rows = rect.r1 - rect.r0 + 1
  const cols = rect.c1 - rect.c0 + 1
  if (rows * cols < 2) return
  await tableMergeOp(ctx, table.sourceId, 'merge-range', rect.r0, rect.c0, { rows, cols }, (id) =>
    singleCellSel(id, { row: rect.r0, col: rect.c0 }),
  )
}

/** Split Cells: undo every merge inside the block */
export async function splitSelectedCells(ctx: ActionCtx): Promise<void> {
  const range = currentRange(ctx)
  if (!range) return
  const { table, rect } = range
  await tableMergeOp(
    ctx,
    table.sourceId,
    'split-range',
    rect.r0,
    rect.c0,
    { rows: rect.r1 - rect.r0 + 1, cols: rect.c1 - rect.c0 + 1 },
    (id) => ({
      sourceId: id,
      anchor: { row: rect.r0, col: rect.c0 },
      focus: { row: rect.r1, col: rect.c1 },
    }),
  )
}

/** Whether a split would do anything (the block holds a merged cell) */
export function rangeHasMerge(table: TableRenderNode, rect: CellRect): boolean {
  return cellAddrsInRect(table, rect).some((a) => {
    const c = table.cells.find((x) => x.row === a.row && x.col === a.col)
    return !!c && ((c.gridSpan ?? 1) > 1 || (c.rowSpan ?? 1) > 1)
  })
}

export type StructureKind =
  'row-above' | 'row-below' | 'col-left' | 'col-right' | 'del-row' | 'del-col'

/**
 * Insert / delete as many rows or columns as the block spans (PowerPoint: with two rows selected,
 * Insert Above adds two rows). The selection keeps pointing at the same cells afterwards.
 */
export async function structureForRange(
  ctx: ActionCtx,
  table: TableRenderNode,
  rect: CellRect,
  kind: StructureKind,
): Promise<void> {
  const rows = rect.r1 - rect.r0 + 1
  const cols = rect.c1 - rect.c0 + 1
  const id = table.sourceId
  const keep =
    (dr: number, dc: number) =>
    (newId: string): TableCellSel => ({
      sourceId: newId,
      anchor: { row: rect.r0 + dr, col: rect.c0 + dc },
      focus: { row: rect.r1 + dr, col: rect.c1 + dc },
    })
  switch (kind) {
    case 'row-above':
      return tableStructureOp(ctx, id, 'insert-row', rect.r0, true, rows, keep(rows, 0))
    case 'row-below':
      return tableStructureOp(ctx, id, 'insert-row', rect.r1, false, rows, keep(0, 0))
    case 'col-left':
      return tableStructureOp(ctx, id, 'insert-col', rect.c0, true, cols, keep(0, cols))
    case 'col-right':
      return tableStructureOp(ctx, id, 'insert-col', rect.c1, false, cols, keep(0, 0))
    case 'del-row':
      return tableStructureOp(ctx, id, 'delete-row', rect.r0, false, rows, (newId) =>
        singleCellSel(newId, {
          row: Math.min(rect.r0, gridRows(table) - rows - 1),
          col: rect.c0,
        }),
      )
    case 'del-col':
      return tableStructureOp(ctx, id, 'delete-col', rect.c0, false, cols, (newId) =>
        singleCellSel(newId, {
          row: rect.r0,
          col: Math.min(rect.c0, gridCols(table) - cols - 1),
        }),
      )
  }
}

/** Enter text editing in one cell (Enter / F2 / double-click / typing onto the selection) */
export function editCell(ctx: ActionCtx, at: CellAddr, replaceWith?: string): void {
  const table = selectedTable(ctx)
  if (!table) return
  ctx.setEditing(null)
  ctx.setCellSel(null)
  ctx.setEditingCell({
    sourceId: table.sourceId,
    row: at.row,
    col: at.col,
    ...(replaceWith ? { replaceWith } : {}),
  })
}

const ARROWS: Record<string, Dir> = {
  ArrowUp: 'up',
  ArrowDown: 'down',
  ArrowLeft: 'left',
  ArrowRight: 'right',
}

/**
 * Keys while cells (not text) are selected. Returns true when the key was consumed.
 * Esc leaves cell mode (the table stays selected); Enter / F2 / typing enter the cell's text.
 */
export function handleCellSelectionKey(ctx: ActionCtx, e: KeyboardEvent): boolean {
  if (ctx.editing || ctx.editingCell || ctx.editingNotes) return false
  const sel = ctx.cellSel
  const table = selectedTable(ctx)
  if (!sel || !table) return false
  const mod = e.metaKey || e.ctrlKey
  const key = e.key

  const dir = ARROWS[key]
  if (dir) {
    // arrows never nudge the table while its cells are selected (⌘/⌥ + arrow do nothing here)
    e.preventDefault()
    if (!mod && !e.altKey) {
      ctx.setCellSel(e.shiftKey ? extendSel(table, sel, dir) : moveSel(table, sel, dir))
    }
    return true
  }
  if (key === 'Tab' && !mod && !e.altKey) {
    e.preventDefault()
    const from = e.shiftKey ? selectionTopLeft(table, sel) : selectionBottomRight(table, sel)
    const next = tabTarget(table, from, e.shiftKey ? -1 : 1)
    if (next) ctx.setCellSel(singleCellSel(sel.sourceId, next))
    else if (!e.shiftKey) {
      // Tab past the last cell grows the table by a row, like PowerPoint
      const lastRow = gridRows(table) - 1
      void tableStructureOp(ctx, sel.sourceId, 'insert-row', lastRow, false, 1, (id) =>
        singleCellSel(id, { row: lastRow + 1, col: 0 }),
      )
    }
    return true
  }
  if ((key === 'Enter' || key === 'F2') && !mod && !e.altKey && !e.shiftKey) {
    e.preventDefault()
    editCell(ctx, sel.anchor)
    return true
  }
  if (key === 'Escape') {
    e.preventDefault()
    ctx.setCellSel(null)
    return true
  }
  if ((key === 'Delete' || key === 'Backspace') && !mod && !e.altKey) {
    e.preventDefault()
    void clearSelectedCells(ctx)
    return true
  }
  if (mod && !e.altKey && !e.shiftKey) {
    const k = key.toLowerCase()
    if (k === 'a') {
      e.preventDefault()
      ctx.setCellSel(allSel(table, sel.sourceId))
      return true
    }
    if (k === 'c' || k === 'x') {
      e.preventDefault()
      void copySelectedCells(ctx, k === 'x')
      return true
    }
    if (k === 'v') {
      e.preventDefault()
      void pasteIntoCells(ctx)
      return true
    }
  }
  // typing a character onto the selection starts editing the first cell with it
  if (!mod && !e.altKey && key.length === 1) {
    e.preventDefault()
    editCell(ctx, sel.anchor, key)
    return true
  }
  return false
}

function selectionTopLeft(table: TableRenderNode, sel: TableCellSel): CellAddr {
  const r = selectionRect(table, sel)
  return { row: r.r0, col: r.c0 }
}

function selectionBottomRight(table: TableRenderNode, sel: TableCellSel): CellAddr {
  // Tab walks on from the last cell of the block
  const r = selectionRect(table, sel)
  const c = cellCovering(table, r.r1, r.c1)
  return c ? { row: c.row, col: c.col } : { row: r.r1, col: r.c1 }
}
