/**
 * Table Layout tab commands that PowerPoint offers beyond the canvas gestures: evenly distributing
 * rows/columns, sizing the whole table, aligning text vertically. They run as one transaction
 * (a single undo step) through the same ops the AI uses.
 */
import type { RenderSlide, TableRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import { MAX_ELEMENT_PX, colWidthFloor, resolveRowDrag, sanitizeSizePx } from './min-size'
import { minSizeOfNodes } from './selection-min'

/** layout px → EMU (96 dpi layout px, scaled by the slide's zoom) */
const emuOf = (px: number, scale: number): number => Math.max(1, Math.round((px / scale) * 9525))

function currentTable(ctx: ActionCtx): { table: TableRenderNode; slide: RenderSlide } | null {
  const { slide, selectedNode } = ctx
  if (!slide || selectedNode?.type !== 'table') return null
  return { table: selectedNode as TableRenderNode, slide }
}

const gaps = (grid: number[]): number[] => grid.slice(1).map((v, i) => v - grid[i]!)

export async function runOps(ctx: ActionCtx, ops: unknown[]): Promise<void> {
  if (!ops.length) return
  const r = await window.slidesApi.applyTxn?.({ ops })
  const next = r?.applied ? r.slides?.[ctx.current] : undefined
  if (next) ctx.applySlide(ctx.current, next)
}

/**
 * Even out row heights / column widths. With a cell range spanning several rows (columns) only
 * those rows (columns) are evened out, like PowerPoint; otherwise the whole table.
 */
export async function distributeTable(
  ctx: ActionCtx,
  axis: 'rows' | 'cols',
  span?: { from: number; to: number } | null,
): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur) return
  const { table, slide } = cur
  const target = { slide: ctx.current, el: table.sourceId }
  const grid = axis === 'rows' ? table.gridY : table.gridX
  const all = gaps(grid)
  const from = span && span.to > span.from ? span.from : 0
  const to = span && span.to > span.from ? span.to : all.length - 1
  const idx = Array.from({ length: to - from + 1 }, (_, i) => from + i)
  const each = emuOf(idx.reduce((a, i) => a + all[i]!, 0) / idx.length, slide.scale)
  await runOps(
    ctx,
    idx.map((i) =>
      axis === 'rows'
        ? { op: 'setTableRowHeight', target, row: i, hEmu: each }
        : { op: 'setTableColWidth', target, col: i, wEmu: each },
    ),
  )
}

/** Cell Size fields: every selected row gets this height / every selected column this width (one undo step). */
export async function setSelectedRowHeights(
  ctx: ActionCtx,
  rows: number[],
  hPx: number,
): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur || !Number.isFinite(hPx)) return
  const target = { slide: ctx.current, el: cur.table.sourceId }
  // a typed height below the floor / above PowerPoint's limit is clamped, never rejected silently
  const h = resolveRowDrag(hPx, hPx) // finite already; clamps to [MIN_CELL_ROW_PX, MAX_ELEMENT_PX]
  await runOps(
    ctx,
    rows.map((row) => ({
      op: 'setTableRowHeight',
      target,
      row,
      hEmu: emuOf(Math.min(h, MAX_ELEMENT_PX * cur.slide.scale), cur.slide.scale),
    })),
  )
}

export async function setSelectedColWidths(
  ctx: ActionCtx,
  cols: number[],
  wPx: number,
): Promise<void> {
  const cur = currentTable(ctx)
  // a typed width stops where the neediest selected column would start cutting a word apart
  const floor = cur ? colWidthFloor(gaps(cur.table.gridX), cols, cur.table.colNeedPx) : 0
  const w = cur ? sanitizeSizePx(wPx, floor, MAX_ELEMENT_PX * cur.slide.scale) : null
  if (!cur || w == null) return
  const target = { slide: ctx.current, el: cur.table.sourceId }
  await runOps(
    ctx,
    cols.map((col) => ({ op: 'setTableColWidth', target, col, wEmu: emuOf(w, cur.slide.scale) })),
  )
}

/** Scale every row (h) or column (w) by the same factor so the table reaches `px`. */
export async function resizeTable(ctx: ActionCtx, axis: 'h' | 'w', px: number): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur) return
  const { table, slide } = cur
  // same floor as dragging the table's handles, same ceiling as the size pane
  const floor = minSizeOfNodes([table])
  const want = sanitizeSizePx(px, axis === 'h' ? floor.h : floor.w, MAX_ELEMENT_PX * slide.scale)
  if (want == null) return
  const target = { slide: ctx.current, el: table.sourceId }
  const sizes = gaps(axis === 'h' ? table.gridY : table.gridX)
  const total = sizes.reduce((a, b) => a + b, 0)
  if (!(total > 0)) return
  const k = want / total
  await runOps(
    ctx,
    sizes.map((s, i) =>
      axis === 'h'
        ? { op: 'setTableRowHeight', target, row: i, hEmu: emuOf(s * k, slide.scale) }
        : { op: 'setTableColWidth', target, col: i, wEmu: emuOf(s * k, slide.scale) },
    ),
  )
}

/** Vertical text alignment of the selected cells (range or the cell being edited), else of every cell. */
export async function alignTableCells(
  ctx: ActionCtx,
  anchor: 'top' | 'middle' | 'bottom',
): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur) return
  const { table } = cur
  const target = { slide: ctx.current, el: table.sourceId }
  const cells = ctx.tableCells?.length
    ? ctx.tableCells
    : table.cells.map((c) => ({ row: c.row, col: c.col }))
  await runOps(
    ctx,
    cells.map((c) => ({ op: 'setTableCellAnchor', target, row: c.row, col: c.col, anchor })),
  )
}
