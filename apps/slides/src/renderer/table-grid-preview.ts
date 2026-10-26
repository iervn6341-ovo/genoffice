/**
 * Pure geometry of the live column / row border drag. While the pointer moves, the table is re-gridded
 * locally (cells, borders, text blocks follow the new grid lines on the very next frame); the
 * authoritative layout — text re-wrapped at the new widths, rows grown to fit — arrives from the main
 * process a moment later and is re-gridded onto the pointer's latest position, so the columns never
 * rubber-band back to a stale pointer position.
 */
import type { TableCellRender, TableRenderNode } from '@genoffice/pptx-render'

/** Column widths of a grid line list */
export function gapsOf(grid: number[]): number[] {
  return grid.slice(1).map((v, i) => v - grid[i]!)
}

/** Grid line offsets of a width list */
export function gridOf(sizes: number[]): number[] {
  const out = [0]
  for (const s of sizes) out.push(out[out.length - 1]! + s)
  return out
}

/** The border drag the pointer asks for, in absolute px */
export type GridChange =
  | { kind: 'col'; col: number; wPx: number; neighbor?: { col: number; wPx: number } }
  | { kind: 'row'; row: number; hPx: number }

/** Grid lines of `tbl` with the dragged border applied (all other columns / rows keep their size) */
export function changedGrid(
  tbl: Pick<TableRenderNode, 'gridX' | 'gridY'>,
  change: GridChange,
): { gridX: number[]; gridY: number[] } {
  if (change.kind === 'col') {
    const w = gapsOf(tbl.gridX)
    if (change.col >= 0 && change.col < w.length) w[change.col] = change.wPx
    if (change.neighbor && change.neighbor.col >= 0 && change.neighbor.col < w.length)
      w[change.neighbor.col] = change.neighbor.wPx
    return { gridX: gridOf(w), gridY: tbl.gridY }
  }
  const h = gapsOf(tbl.gridY)
  if (change.row >= 0 && change.row < h.length) h[change.row] = change.hPx
  return { gridX: tbl.gridX, gridY: gridOf(h) }
}

/** Piecewise-linear map from the old knot positions to the new ones (shifts beyond the ends) */
function knotMap(from: number[], to: number[]): (v: number) => number {
  const n = Math.min(from.length, to.length)
  return (v) => {
    if (n < 2) return v
    if (v <= from[0]!) return to[0]! + (v - from[0]!)
    for (let i = 0; i < n - 1; i++) {
      const a = from[i]!
      const b = from[i + 1]!
      if (v <= b + 1e-9) {
        if (b - a < 1e-9) return to[i]!
        return to[i]! + ((v - a) / (b - a)) * (to[i + 1]! - to[i]!)
      }
    }
    return to[n - 1]! + (v - from[n - 1]!)
  }
}

/** Visual (left→right) knot positions of the column grid; rtl tables mirror the logical grid */
function visualKnots(grid: number[], rtl: boolean | undefined): number[] {
  const tw = grid[grid.length - 1] ?? 0
  return rtl ? grid.map((v) => tw - v).reverse() : grid
}

const ANCHOR_SHARE = { top: 0, middle: 0.5, bottom: 1 } as const

/**
 * `tbl` re-gridded onto new grid lines: every cell edge is mapped through the old→new line positions
 * (merged cells stretch across the lines they span), vertically anchored text keeps its place inside
 * the cell, and the node box follows the new total size. Row heights that grew with re-wrapped text
 * are the caller's business (it passes the grown grid) — nothing is re-measured here.
 */
export function regridTable(
  tbl: TableRenderNode,
  gridX: number[],
  gridY: number[],
): TableRenderNode {
  if (gridX.length !== tbl.gridX.length || gridY.length !== tbl.gridY.length) return tbl
  const fx = knotMap(visualKnots(tbl.gridX, tbl.rtl), visualKnots(gridX, tbl.rtl))
  const fy = knotMap(tbl.gridY, gridY)
  const cells: TableCellRender[] = tbl.cells.map((c) => {
    const x = fx(c.x)
    const y = fy(c.y)
    const w = fx(c.x + c.w) - x
    const h = fy(c.y + c.h) - y
    const out: TableCellRender = { ...c, x, y, w, h }
    const share = c.text && !c.text.vert ? ANCHOR_SHARE[c.text.anchor] : 0
    if (c.text && share && Math.abs(h - c.h) > 1e-6) {
      out.text = {
        ...c.text,
        insets: { ...c.text.insets, t: c.text.insets.t + (h - c.h) * share },
      }
    }
    return out
  })
  return {
    ...tbl,
    cells,
    gridX,
    gridY,
    // the frame may be larger than the grid (stale ext): keep that slack, move only by the grid's growth
    box: {
      ...tbl.box,
      w: tbl.box.w + ((gridX[gridX.length - 1] ?? 0) - (tbl.gridX[tbl.gridX.length - 1] ?? 0)),
      h: tbl.box.h + ((gridY[gridY.length - 1] ?? 0) - (tbl.gridY[tbl.gridY.length - 1] ?? 0)),
    },
  }
}
