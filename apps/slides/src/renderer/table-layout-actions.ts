/**
 * Table Layout tab commands that PowerPoint offers beyond the canvas gestures: evenly distributing
 * rows/columns, sizing the whole table, aligning text vertically. They run as one transaction
 * (a single undo step) through the same ops the AI uses.
 */
import type { RenderSlide, TableRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'

/** layout px → EMU (96 dpi layout px, scaled by the slide's zoom) */
const emuOf = (px: number, scale: number): number => Math.max(1, Math.round((px / scale) * 9525))

function currentTable(ctx: ActionCtx): { table: TableRenderNode; slide: RenderSlide } | null {
  const { slide, selectedNode } = ctx
  if (!slide || selectedNode?.type !== 'table') return null
  return { table: selectedNode as TableRenderNode, slide }
}

const gaps = (grid: number[]): number[] => grid.slice(1).map((v, i) => v - grid[i]!)

async function runOps(ctx: ActionCtx, ops: unknown[]): Promise<void> {
  if (!ops.length) return
  const r = await window.slidesApi.applyTxn?.({ ops })
  const next = r?.applied ? r.slides?.[ctx.current] : undefined
  if (next) ctx.applySlide(ctx.current, next)
}

export async function distributeTable(ctx: ActionCtx, axis: 'rows' | 'cols'): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur) return
  const { table, slide } = cur
  const target = { slide: ctx.current, el: table.sourceId }
  if (axis === 'rows') {
    const sizes = gaps(table.gridY)
    const each = emuOf(sizes.reduce((a, b) => a + b, 0) / sizes.length, slide.scale)
    await runOps(
      ctx,
      sizes.map((_, row) => ({ op: 'setTableRowHeight', target, row, hEmu: each })),
    )
  } else {
    const sizes = gaps(table.gridX)
    const each = emuOf(sizes.reduce((a, b) => a + b, 0) / sizes.length, slide.scale)
    await runOps(
      ctx,
      sizes.map((_, col) => ({ op: 'setTableColWidth', target, col, wEmu: each })),
    )
  }
}

/** Scale every row (h) or column (w) by the same factor so the table reaches `px`. */
export async function resizeTable(ctx: ActionCtx, axis: 'h' | 'w', px: number): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur || !(px > 0)) return
  const { table, slide } = cur
  const target = { slide: ctx.current, el: table.sourceId }
  const sizes = gaps(axis === 'h' ? table.gridY : table.gridX)
  const total = sizes.reduce((a, b) => a + b, 0)
  if (!(total > 0)) return
  const k = px / total
  await runOps(
    ctx,
    sizes.map((s, i) =>
      axis === 'h'
        ? { op: 'setTableRowHeight', target, row: i, hEmu: emuOf(s * k, slide.scale) }
        : { op: 'setTableColWidth', target, col: i, wEmu: emuOf(s * k, slide.scale) },
    ),
  )
}

/** Vertical text alignment of the cell being edited, or of every cell when there is no cell cursor. */
export async function alignTableCells(
  ctx: ActionCtx,
  anchor: 'top' | 'middle' | 'bottom',
): Promise<void> {
  const cur = currentTable(ctx)
  if (!cur) return
  const { table } = cur
  const target = { slide: ctx.current, el: table.sourceId }
  const cells = ctx.editingCell
    ? [{ row: ctx.editingCell.row, col: ctx.editingCell.col }]
    : table.cells.map((c) => ({ row: c.row, col: c.col }))
  await runOps(
    ctx,
    cells.map((c) => ({ op: 'setTableCellAnchor', target, row: c.row, col: c.col, anchor })),
  )
}
