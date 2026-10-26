import { describe, expect, it } from 'vitest'
import type { TableCellRender, TableRenderNode } from '@genoffice/pptx-render'
import {
  changedGrid,
  gapsOf,
  gridOf,
  regridTable,
  type GridChange,
} from '../src/renderer/table-grid-preview'

function cell(
  row: number,
  col: number,
  x: number,
  y: number,
  w: number,
  h: number,
  extra: Partial<TableCellRender> = {},
): TableCellRender {
  return { row, col, x, y, w, h, fill: { kind: 'none' }, ...extra } as TableCellRender
}

/** 3 columns (100/200/100) x 2 rows (40/60) */
function table(extra: Partial<TableRenderNode> = {}): TableRenderNode {
  const gridX = [0, 100, 300, 400]
  const gridY = [0, 40, 100]
  const cells: TableCellRender[] = []
  for (let r = 0; r < 2; r++)
    for (let c = 0; c < 3; c++)
      cells.push(
        cell(r, c, gridX[c]!, gridY[r]!, gridX[c + 1]! - gridX[c]!, gridY[r + 1]! - gridY[r]!),
      )
  return {
    id: 'r_t',
    type: 'table',
    sourceId: 't',
    box: { x: 50, y: 60, w: 400, h: 100, rotationDeg: 0 },
    cells,
    gridX,
    gridY,
    ...extra,
  } as unknown as TableRenderNode
}

describe('grid helpers', () => {
  it('gapsOf / gridOf round-trip', () => {
    expect(gapsOf([0, 100, 300, 400])).toEqual([100, 200, 100])
    expect(gridOf([100, 200, 100])).toEqual([0, 100, 300, 400])
  })

  it('an inner column border moves between its two columns', () => {
    const change: GridChange = { kind: 'col', col: 0, wPx: 60, neighbor: { col: 1, wPx: 240 } }
    const g = changedGrid(table(), change)
    expect(g.gridX).toEqual([0, 60, 300, 400])
    expect(g.gridY).toEqual([0, 40, 100])
  })

  it('the outer column border changes the table width', () => {
    expect(changedGrid(table(), { kind: 'col', col: 2, wPx: 150 }).gridX).toEqual([
      0, 100, 300, 450,
    ])
  })

  it('a row border changes that row only', () => {
    expect(changedGrid(table(), { kind: 'row', row: 0, hPx: 70 }).gridY).toEqual([0, 70, 130])
  })
})

describe('regridTable', () => {
  it('moves every cell edge with its grid line', () => {
    const t = table()
    const g = changedGrid(t, { kind: 'col', col: 0, wPx: 60, neighbor: { col: 1, wPx: 240 } })
    const out = regridTable(t, g.gridX, g.gridY)
    const c01 = out.cells.find((c) => c.row === 0 && c.col === 1)!
    expect([c01.x, c01.w]).toEqual([60, 240])
    const c00 = out.cells.find((c) => c.row === 0 && c.col === 0)!
    expect([c00.x, c00.w]).toEqual([0, 60])
    // the third column did not move
    const c02 = out.cells.find((c) => c.row === 1 && c.col === 2)!
    expect([c02.x, c02.w]).toEqual([300, 100])
    expect(out.box.w).toBe(400)
    // the source node is untouched
    expect(t.gridX).toEqual([0, 100, 300, 400])
  })

  it('a merged cell stretches across the lines it spans', () => {
    const t = table()
    t.cells = [cell(0, 0, 0, 0, 300, 40, { gridSpan: 2 }), ...t.cells.filter((c) => c.col === 2)]
    const g = changedGrid(t, { kind: 'col', col: 0, wPx: 40, neighbor: { col: 1, wPx: 260 } })
    const out = regridTable(t, g.gridX, g.gridY)
    const merged = out.cells.find((c) => c.gridSpan === 2)!
    // both of its columns together keep their total: it does not change size
    expect([merged.x, merged.w]).toEqual([0, 300])
    const wide = changedGrid(t, { kind: 'col', col: 1, wPx: 300 })
    const out2 = regridTable(t, wide.gridX, wide.gridY)
    expect(out2.cells.find((c) => c.gridSpan === 2)!.w).toBe(400)
    expect(out2.box.w).toBe(500)
  })

  it('the node box grows by the grid growth, keeping a stale-frame slack', () => {
    const t = table({ box: { x: 50, y: 60, w: 400, h: 180, rotationDeg: 0 } } as never)
    const g = changedGrid(t, { kind: 'row', row: 1, hPx: 90 })
    const out = regridTable(t, g.gridX, g.gridY)
    expect(out.box.h).toBe(210)
    expect(out.box.x).toBe(50)
  })

  it('middle- and bottom-anchored text keeps its place in a taller cell', () => {
    const t = table()
    const text = (anchor: 'top' | 'middle' | 'bottom') =>
      ({ anchor, insets: { l: 5, t: 4, r: 5, b: 4 }, lines: [], contentHeight: 10 }) as never
    t.cells = [
      cell(0, 0, 0, 0, 100, 40, { text: text('top') }),
      cell(0, 1, 100, 0, 200, 40, { text: text('middle') }),
      cell(0, 2, 300, 0, 100, 40, { text: text('bottom') }),
    ]
    const g = changedGrid(t, { kind: 'row', row: 0, hPx: 80 })
    const out = regridTable(t, g.gridX, g.gridY)
    expect(out.cells[0]!.text!.insets.t).toBe(4)
    expect(out.cells[1]!.text!.insets.t).toBe(4 + 20)
    expect(out.cells[2]!.text!.insets.t).toBe(4 + 40)
  })

  it('rtl tables are mirrored: the logical grid grows leftward, visual cells follow', () => {
    // logical columns 100/200/100, rendered mirrored: logical col 0 is rightmost
    const t = table({ rtl: true } as never)
    t.cells = [cell(0, 0, 300, 0, 100, 40), cell(0, 1, 100, 0, 200, 40), cell(0, 2, 0, 0, 100, 40)]
    const g = changedGrid(t, { kind: 'col', col: 2, wPx: 150 })
    const out = regridTable(t, g.gridX, g.gridY)
    // logical col 2 is the visual leftmost: it widens to 150, everything else shifts right by 50
    const byCol = (c: number) => out.cells.find((x) => x.col === c)!
    expect([byCol(2).x, byCol(2).w]).toEqual([0, 150])
    expect([byCol(1).x, byCol(1).w]).toEqual([150, 200])
    expect([byCol(0).x, byCol(0).w]).toEqual([350, 100])
  })

  it('a grid of another shape is refused (the node comes back as it was)', () => {
    const t = table()
    expect(regridTable(t, [0, 10], t.gridY)).toBe(t)
  })
})
