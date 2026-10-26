import { describe, expect, it } from 'vitest'
import type { RenderNode } from '@genoffice/pptx-render'
import { minSizeOfNodes, toMinNode } from '../src/renderer/selection-min'

const table = (extra: Record<string, unknown> = {}): RenderNode =>
  ({
    id: 'r_t',
    type: 'table',
    sourceId: 't',
    box: { x: 0, y: 0, w: 560, h: 100, rotationDeg: 0, flipH: false, flipV: false },
    cells: [],
    gridX: [0, 140, 400, 560],
    gridY: [0, 50, 100],
    rowNeedPx: [30, 45],
    colNeedPx: [60, 70, 112],
    ...extra,
  }) as unknown as RenderNode

describe('table selection floor', () => {
  it('carries the measured column needs and the bottom-fit width into the floor', () => {
    const mn = toMinNode(table({ fitMinWPx: 430 }))
    expect(mn.colNeeds).toEqual([60, 70, 112])
    expect(mn.fitMinW).toBe(430)
  })

  it('a render node without needs keeps the legacy per-column floor', () => {
    expect(minSizeOfNodes([table({ colNeedPx: undefined })]).w).toBe(3 * 64)
  })

  it('narrows to the first column that would cut a word, not to 64px per column', () => {
    const m = minSizeOfNodes([table()])
    expect(m.w).toBeCloseTo(560 * (112.5 / 160), 6)
    expect(m.w).toBeLessThan(560)
    expect(m.h).toBe(30 + 45)
  })

  it('the bottom-fit width wins when it is the larger floor', () => {
    expect(minSizeOfNodes([table({ fitMinWPx: 450 })]).w).toBe(450)
  })
})
