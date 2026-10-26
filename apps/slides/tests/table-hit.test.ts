import type { TableRenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import {
  TABLE_BORDER_GRIP_PX,
  tableCellAtPoint,
  tableCellInterior,
  tableCellOverlayBox,
  tableLocalPointFromStage,
} from '../src/renderer/table-hit'

function makeTable(): TableRenderNode {
  return {
    type: 'table',
    sourceId: 't1',
    id: 'node_t1',
    box: { x: 100, y: 50, w: 120, h: 80, rotationDeg: 0 },
    cells: [
      { x: 0, y: 0, w: 60, h: 40, row: 0, col: 0, fill: { kind: 'solid', color: '#fff' } },
      { x: 60, y: 0, w: 60, h: 40, row: 0, col: 1, fill: { kind: 'solid', color: '#fff' } },
      { x: 0, y: 40, w: 60, h: 40, row: 1, col: 0, fill: { kind: 'solid', color: '#fff' } },
      { x: 60, y: 40, w: 60, h: 40, row: 1, col: 1, fill: { kind: 'solid', color: '#fff' } },
    ],
    gridX: [0, 60, 120],
    gridY: [0, 40, 80],
  } as unknown as TableRenderNode
}

describe('table cell hit testing', () => {
  it('resolves a cell in unrotated table coordinates', () => {
    const table = makeTable()
    expect(tableCellAtPoint(table, { x: 20, y: 10 })).toMatchObject({ row: 0, col: 0 })
    expect(tableCellAtPoint(table, { x: 95, y: 12 })).toMatchObject({ row: 0, col: 1 })
  })

  it('rotates stage hits back into table-local coordinates', () => {
    const table = makeTable()
    const rad = (45 * Math.PI) / 180
    const local = { x: 90, y: 60 }
    const stagePoint = {
      x: 160 + (local.x - 60) * Math.cos(rad) - (local.y - 40) * Math.sin(rad),
      y: 90 + (local.x - 60) * Math.sin(rad) + (local.y - 40) * Math.cos(rad),
    }
    const inverse = tableLocalPointFromStage(stagePoint, {
      x: table.box.x,
      y: table.box.y,
      w: table.box.w,
      h: table.box.h,
      rotationDeg: 45,
    })
    expect(inverse).toMatchObject({ x: 90, y: 60 })
    expect(tableCellAtPoint(table, inverse)).toMatchObject({ row: 1, col: 1 })
  })

  it('places the edit overlay at the rotated cell center', () => {
    const overlay = tableCellOverlayBox(
      { x: 100, y: 50, w: 120, h: 80, rotationDeg: 90 },
      { x: 60, y: 40, w: 60, h: 40 },
    )
    expect(overlay).toMatchObject({ x: 110, y: 100, w: 60, h: 40, rotationDeg: 90 })
  })
})

describe('tableCellInterior (cell body vs. border grip)', () => {
  const gap = TABLE_BORDER_GRIP_PX

  it('returns the cell for a press well inside it', () => {
    expect(tableCellInterior(makeTable(), { x: 30, y: 20 }, gap)).toMatchObject({ row: 0, col: 0 })
    expect(tableCellInterior(makeTable(), { x: 90, y: 60 }, gap)).toMatchObject({ row: 1, col: 1 })
  })

  it('treats a press on a cell border or the table frame as the move grip (no cell)', () => {
    const t = makeTable()
    expect(tableCellInterior(t, { x: 60, y: 20 }, gap)).toBeUndefined() // shared column border
    expect(tableCellInterior(t, { x: 58, y: 20 }, gap)).toBeUndefined() // within the band
    expect(tableCellInterior(t, { x: 30, y: 40 }, gap)).toBeUndefined() // shared row border
    expect(tableCellInterior(t, { x: 1, y: 20 }, gap)).toBeUndefined() // outer frame
    expect(tableCellInterior(t, { x: 30, y: 79 }, gap)).toBeUndefined()
  })

  it('returns nothing outside the table', () => {
    expect(tableCellInterior(makeTable(), { x: 500, y: 500 }, gap)).toBeUndefined()
  })

  it('scales the band with the gap argument (zoomed-out canvas = larger slide-px band)', () => {
    const t = makeTable()
    expect(tableCellInterior(t, { x: 10, y: 20 }, 5)).toBeDefined()
    expect(tableCellInterior(t, { x: 10, y: 20 }, 12)).toBeUndefined()
  })
})
