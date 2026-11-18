import { expect, it, vi } from 'vitest'
import { createSelectionDrag, selectionMoveRequests } from '../src/renderer/selection-drag'
import { createTransformBatcher, type TransformReq } from '../src/renderer/history-group'
import type { RenderNode } from '@genoffice/pptx-render'

it('uses one snapped pointer displacement for every member despite proxy-drag offsets', () => {
  const drag = createSelectionDrag()
  const snap = vi.fn(({ x, y }) => ({ x: Math.round(x / 10) * 10, y }))
  drag.begin(
    { x: 50, y: 70 },
    new Map([
      ['shape', { x: 100, y: 110 }],
      ['group', { x: 210, y: 230 }],
      ['table', { x: 340, y: 200 }],
    ]),
    snap,
  )
  for (const pointer of [
    { x: 58, y: 76 },
    { x: 62, y: 90 },
    { x: 77, y: 95 },
  ]) {
    const a = drag.bound('shape', pointer)!
    const b = drag.bound('group', pointer)!
    const c = drag.bound('table', pointer)!
    expect({ x: b.x - a.x, y: b.y - a.y }).toEqual({ x: 110, y: 120 })
    expect({ x: c.x - a.x, y: c.y - a.y }).toEqual({ x: 240, y: 90 })
  }
  expect(snap).toHaveBeenCalledTimes(3)
})

it('captures positions before proxy nodes move and discards the previous gesture on a new press', () => {
  const drag = createSelectionDrag()
  const positions = new Map([['a', { x: 3, y: 4 }]])
  drag.begin({ x: 0, y: 0 }, positions, (p) => p)
  positions.get('a')!.x = 999
  expect(drag.bound('a', { x: 10, y: 20 })).toEqual({ x: 13, y: 24 })
  drag.reset()
  expect(drag.bound('a', { x: 10, y: 20 })).toBeNull()
  drag.begin({ x: 10, y: 20 }, new Map([['a', { x: 13, y: 24 }]]), (p) => p)
  expect(drag.bound('a', { x: 15, y: 15 })).toEqual({ x: 18, y: 19 })
  expect(drag.bound('unselected', { x: 15, y: 15 })).toBeNull()
})

it('one leader dragend commits every selected member in a single batch and only once', async () => {
  const nodes = [0, 1, 2].map((i) => ({
    sourceId: `n${i}`,
    box: { x: 50 + i * 100, y: 80, w: 90, h: 40, rotationDeg: i * 25 },
  })) as RenderNode[]
  const drag = createSelectionDrag()
  drag.begin(
    { x: 0, y: 0 },
    new Map(nodes.map((n) => [n.sourceId, { x: n.box.x, y: n.box.y }])),
    (d) => d,
  )
  drag.bound('n0', { x: 12, y: 34 }) // Konva may emit dragend only for this node
  const delta = drag.finish()!
  expect(drag.finish()).toBeNull()
  const flush = vi.fn(async (_slide: number, _reqs: TransformReq[]) => {})
  const scheduled: Array<() => void> = []
  const batch = createTransformBatcher(flush, (fn) => scheduled.push(fn))
  const commits = selectionMoveRequests(nodes, delta).map((r) => batch.push(2, r))
  scheduled.forEach((fn) => fn())
  await Promise.all(commits)
  expect(flush).toHaveBeenCalledTimes(1)
  expect(flush.mock.calls[0]![1]).toEqual(
    nodes.map((n) => ({
      sourceId: n.sourceId,
      box: { ...n.box, x: n.box.x + 12, y: n.box.y + 34 },
    })),
  )
})

it('converts a group-child move into its parent coordinates without changing shape sizes or rotation', () => {
  const nodes = [
    { sourceId: 'child', box: { x: 40, y: 50, w: 60, h: 70, rotationDeg: 25 } },
  ] as RenderNode[]
  const [req] = selectionMoveRequests(nodes, { x: 20, y: 30 }, 'group', (_id, d) => ({
    x: d.y,
    y: -d.x,
  }))
  expect(req).toEqual({
    sourceId: 'child',
    groupId: 'group',
    box: { x: 70, y: 30, w: 60, h: 70, rotationDeg: 25 },
  })
})

it('moves all 4 elements by exact 20px with zero relative drift across continuous pointer updates', () => {
  const nodes = [0, 1, 2, 3].map((i) => ({
    sourceId: `elem_${i}`,
    box: { x: 50 + i * 120, y: 100, w: 80, h: 60, rotationDeg: 0 },
  })) as RenderNode[]
  const drag = createSelectionDrag()
  const initialPositions = new Map(nodes.map((n) => [n.sourceId, { x: n.box.x, y: n.box.y }]))
  drag.begin({ x: 100, y: 100 }, initialPositions, (d) => d)

  // Move pointer by +20px in X direction
  const pointer = { x: 120, y: 100 }
  const movedPositions = nodes.map((n) => drag.bound(n.sourceId, pointer)!)

  // Every single node must have shifted by exactly +20px in X, 0px in Y
  nodes.forEach((n, idx) => {
    const moved = movedPositions[idx]!
    expect(moved.x - n.box.x).toBe(20)
    expect(moved.y - n.box.y).toBe(0)
  })

  // Distance between consecutive nodes must remain identically 120px
  for (let i = 0; i < movedPositions.length - 1; i++) {
    expect(movedPositions[i + 1]!.x - movedPositions[i]!.x).toBe(120)
  }

  const finalDelta = drag.finish()!
  expect(finalDelta).toEqual({ x: 20, y: 0 })
  const requests = selectionMoveRequests(nodes, finalDelta)
  expect(requests.map((r) => r.box.x)).toEqual([70, 190, 310, 430])
})
