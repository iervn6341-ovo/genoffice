/**
 * Z-order of a multi-selection: the selected elements move together and keep their relative
 * stacking (PowerPoint). Last in the node list = topmost.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { reorderSelected, stackSorted } from '../src/renderer/arrange-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import type { ReorderDirection } from '../src/shared/ipc'

vi.mock('../src/renderer/ink', () => ({ inkNodesOf: () => [], rasterizeStroke: () => ({}) }))

let list: string[] = []
function move(id: string, dir: ReorderDirection) {
  const i = list.indexOf(id)
  list.splice(i, 1)
  const to =
    dir === 'front'
      ? list.length
      : dir === 'back'
        ? 0
        : dir === 'forward'
          ? Math.min(list.length, i + 1)
          : Math.max(0, i - 1)
  list.splice(to, 0, id)
}

function install() {
  ;(globalThis as unknown as { window: unknown }).window = {
    slidesApi: {
      beginHistoryBatch: async () => true,
      endHistoryBatch: async () => undefined,
      reorderElement: async (op: { sourceId: string; dir: ReorderDirection }) => {
        move(op.sourceId, op.dir)
        return { nodes: list.map((id) => ({ sourceId: id })) }
      },
    },
  }
}

const ctxFor = (selected: string[]) =>
  ({
    current: 0,
    selectedIds: selected,
    slide: { nodes: list.map((id) => ({ sourceId: id })) },
    applySlide: vi.fn(),
  }) as unknown as ActionCtx

describe('reorderSelected with several elements', () => {
  beforeEach(() => {
    list = ['a', 'b', 'c', 'd', 'e']
    install()
  })

  it('front keeps the selected pair in their relative order on top', async () => {
    await reorderSelected(ctxFor(['b', 'd']), 'b', 'front')
    expect(list).toEqual(['a', 'c', 'e', 'b', 'd'])
  })

  it('back keeps the selected pair in their relative order at the bottom', async () => {
    await reorderSelected(ctxFor(['b', 'd']), 'd', 'back')
    expect(list).toEqual(['b', 'd', 'a', 'c', 'e'])
  })

  it('forward moves each selected element one step, relative order intact', async () => {
    await reorderSelected(ctxFor(['a', 'b']), 'a', 'forward')
    expect(list).toEqual(['c', 'a', 'b', 'd', 'e'])
  })

  it('backward moves each selected element one step down, relative order intact', async () => {
    await reorderSelected(ctxFor(['d', 'e']), 'e', 'backward')
    expect(list).toEqual(['a', 'b', 'd', 'e', 'c'])
  })

  it('a click target outside the selection reorders just that element', async () => {
    await reorderSelected(ctxFor(['a', 'b']), 'e', 'back')
    expect(list).toEqual(['e', 'a', 'b', 'c', 'd'])
  })
})

describe('stackSorted', () => {
  const nodes = ['a', 'b', 'c', 'd'].map((sourceId) => ({ sourceId }))
  it('orders the moves per direction', () => {
    expect(stackSorted(nodes, ['d', 'a'], 'front')).toEqual(['a', 'd'])
    expect(stackSorted(nodes, ['a', 'd'], 'back')).toEqual(['d', 'a'])
    expect(stackSorted(nodes, ['a', 'd'], 'forward')).toEqual(['d', 'a'])
    expect(stackSorted(nodes, ['d', 'a'], 'backward')).toEqual(['a', 'd'])
  })
})
