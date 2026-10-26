/**
 * Undo granularity: one user action over a whole selection is ONE undo step (PowerPoint).
 * Arrow-key nudge / multi-drag commits go through applyTransformBatch (main) after the
 * renderer's per-tick batcher; grouped multi-call actions collapse through a history group.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import { buildRenderSlide } from '@genoffice/pptx-render'
import type { Session } from '../src/main/session-state'
import {
  beginHistoryBatch,
  endHistoryBatch,
  pushHistory,
  redoStep,
  undoStep,
} from '../src/main/session-state'
import { applyTransformBatch } from '../src/main/batch-transform'
import {
  createTransformBatcher,
  flushTransforms,
  withHistoryGroup,
} from '../src/renderer/history-group'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
  webContents: { fromId: () => null },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

const FIT = 1280

/** A blank deck with a text box, a rectangle and a table on slide 0 (the "marquee over tables and text" case) */
async function sessionWithShapes(): Promise<{ session: Session; ids: string[] }> {
  const opened = await openPptx(await createBlankPptx())
  const session = { path: '', opened, fitWidthPx: FIT, undoStack: [], redoStack: [] } as Session
  const emu = (px: number) => px * 9525
  const r = runTxn(opened, {
    ops: [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: emu(60), y: emu(60), cx: emu(300), cy: emu(60) },
        paragraphs: [{ runs: [{ text: 'Hello' }] }],
      },
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'rect',
        offset: { x: emu(420), y: emu(80), cx: emu(200), cy: emu(120) },
        fill: '#C43E1C',
      },
      {
        op: 'addTable',
        target: { slide: 0 },
        rows: 2,
        cols: 2,
        offset: { x: emu(100), y: emu(300), cx: emu(400), cy: emu(100) },
      },
    ],
  })
  expect(r.applied).toBe(true)
  const ids = buildRenderSlide(opened.deck.slides[0]!, opened.deck.size, { fitWidthPx: FIT })
    .nodes.filter((n) => !n.decoration && ['text', 'shape', 'table'].includes(n.type))
    .map((n) => n.sourceId)
  expect(ids.length).toBe(3)
  ;(session as unknown as { slideIndex: number }).slideIndex = 0
  return { session, ids }
}

function boxes(session: Session, ids: string[]): Array<{ x: number; y: number }> {
  const index = (session as unknown as { slideIndex: number }).slideIndex
  const slide = session.opened.deck.slides[index]!
  const rendered = buildRenderSlide(slide, session.opened.deck.size, { fitWidthPx: FIT })
  return ids.map((id) => {
    const n = rendered.nodes.find((x) => x.sourceId === id)!
    return { x: n.box.x, y: n.box.y }
  })
}

function nudgeItems(session: Session, ids: string[], dx: number, dy: number) {
  const index = (session as unknown as { slideIndex: number }).slideIndex
  const slide = session.opened.deck.slides[index]!
  const rendered = buildRenderSlide(slide, session.opened.deck.size, { fitWidthPx: FIT })
  return ids.map((id) => {
    const b = rendered.nodes.find((x) => x.sourceId === id)!.box
    return {
      sourceId: id,
      xPx: b.x + dx,
      yPx: b.y + dy,
      wPx: b.w,
      hPx: b.h,
      rotationDeg: b.rotationDeg ?? 0,
    }
  })
}

describe('multi-selection nudge is one undo step', () => {
  it('arrow-key move of three elements: one Undo restores all, one Redo moves all again', async () => {
    const { session, ids } = await sessionWithShapes()
    const index = (session as unknown as { slideIndex: number }).slideIndex
    const before = boxes(session, ids)
    expect(
      applyTransformBatch(session, {
        slideIndex: index,
        fitWidthPx: FIT,
        items: nudgeItems(session, ids, 10, 0),
      }),
    ).toBe(true)
    const after = boxes(session, ids)
    after.forEach((b, i) => {
      expect(b.x).toBeCloseTo(before[i]!.x + 10, 0)
      expect(b.y).toBeCloseTo(before[i]!.y, 0)
    })
    expect(session.undoStack.length).toBe(1)

    expect(undoStep(session)).toBe(true)
    boxes(session, ids).forEach((b, i) => {
      expect(b.x).toBeCloseTo(before[i]!.x, 0)
      expect(b.y).toBeCloseTo(before[i]!.y, 0)
    })
    expect(session.undoStack.length).toBe(0)

    expect(redoStep(session)).toBe(true)
    boxes(session, ids).forEach((b, i) => {
      expect(b.x).toBeCloseTo(after[i]!.x, 0)
    })
  })

  it('two presses are two steps, each undoing all three elements', async () => {
    const { session, ids } = await sessionWithShapes()
    const index = (session as unknown as { slideIndex: number }).slideIndex
    const before = boxes(session, ids)
    for (let k = 0; k < 2; k++) {
      applyTransformBatch(session, {
        slideIndex: index,
        fitWidthPx: FIT,
        items: nudgeItems(session, ids, 0, 5),
      })
    }
    expect(session.undoStack.length).toBe(2)
    undoStep(session)
    boxes(session, ids).forEach((b, i) => expect(b.y).toBeCloseTo(before[i]!.y + 5, 0))
    undoStep(session)
    boxes(session, ids).forEach((b, i) => expect(b.y).toBeCloseTo(before[i]!.y, 0))
  })

  it('an invalid item refuses the whole batch and leaves history alone', async () => {
    const { session, ids } = await sessionWithShapes()
    const index = (session as unknown as { slideIndex: number }).slideIndex
    session.redoStack.push({} as never)
    const items = nudgeItems(session, ids, 10, 0)
    items.push({ sourceId: 'no-such-element', xPx: 0, yPx: 0, wPx: 10, hPx: 10, rotationDeg: 0 })
    expect(applyTransformBatch(session, { slideIndex: index, fitWidthPx: FIT, items })).toBe(false)
    expect(session.undoStack.length).toBe(0)
    expect(session.redoStack.length).toBe(1)
  })

  it('the commit after live-preview frames reuses the gesture snapshot (no second step)', async () => {
    const { session, ids } = await sessionWithShapes()
    const index = (session as unknown as { slideIndex: number }).slideIndex
    pushHistory(session) // what the first preview frame does
    session.transformPreview = true
    applyTransformBatch(session, {
      slideIndex: index,
      fitWidthPx: FIT,
      items: nudgeItems(session, ids, 20, 0),
    })
    expect(session.transformPreview).toBe(false)
    expect(session.undoStack.length).toBe(1)
  })
})

describe('history group collapses several edits into one step', () => {
  it('three separate edits inside a group undo together', async () => {
    const { session, ids } = await sessionWithShapes()
    const index = (session as unknown as { slideIndex: number }).slideIndex
    const before = boxes(session, ids)
    beginHistoryBatch(session)
    for (const id of ids) {
      // each per-element handler pushes its own snapshot
      applyTransformBatch(session, {
        slideIndex: index,
        fitWidthPx: FIT,
        items: nudgeItems(session, [id], 7, 0),
      })
    }
    expect(session.undoStack.length).toBe(3)
    endHistoryBatch(session)
    expect(session.undoStack.length).toBe(1)
    expect(undoStep(session)).toBe(true)
    boxes(session, ids).forEach((b, i) => expect(b.x).toBeCloseTo(before[i]!.x, 0))
  })
})

describe('renderer transform batcher', () => {
  it('collects the commits of one tick into one flush and keeps the last box per element', async () => {
    const calls: Array<{ slide: number; ids: string[] }> = []
    let fire: (() => void) | null = null
    const batcher = createTransformBatcher(
      async (slide, reqs) => {
        calls.push({ slide, ids: reqs.map((r) => `${r.sourceId}:${r.box.x}`) })
      },
      (cb) => (fire = cb),
    )
    const box = (x: number) => ({ x, y: 0, w: 10, h: 10, rotationDeg: 0 })
    const p = Promise.all([
      batcher.push(2, { sourceId: 'a', box: box(1) }),
      batcher.push(2, { sourceId: 'b', box: box(2) }),
      batcher.push(2, { sourceId: 'c', box: box(3) }),
      batcher.push(2, { sourceId: 'a', box: box(4) }),
    ])
    expect(calls.length).toBe(0)
    fire!()
    await p
    expect(calls).toEqual([{ slide: 2, ids: ['a:4', 'b:2', 'c:3'] }])
  })

  it('a second tick is a second flush', async () => {
    const calls: number[] = []
    const fires: Array<() => void> = []
    const batcher = createTransformBatcher(
      async (_s, reqs) => void calls.push(reqs.length),
      (cb) => void fires.push(cb),
    )
    const box = { x: 0, y: 0, w: 1, h: 1, rotationDeg: 0 }
    const p1 = batcher.push(0, { sourceId: 'a', box })
    fires.shift()!()
    await p1
    const p2 = Promise.all([
      batcher.push(0, { sourceId: 'a', box }),
      batcher.push(0, { sourceId: 'b', box }),
    ])
    fires.shift()!()
    await p2
    expect(calls).toEqual([1, 2])
  })
})

describe('flushTransforms picks the single-undo path', () => {
  const api = {
    editTransform: vi.fn(),
    batchEditTransform: vi.fn(),
    beginHistoryBatch: vi.fn(),
    endHistoryBatch: vi.fn(),
  }
  const slideResult = { widthPx: 1 } as never
  const box = { x: 1, y: 2, w: 3, h: 4, rotationDeg: 0 }
  beforeEach(() => {
    Object.values(api).forEach((f) => f.mockReset())
    api.beginHistoryBatch.mockResolvedValue(true)
    ;(globalThis as unknown as { window: unknown }).window = { slidesApi: api }
  })

  it('one element → editTransform', async () => {
    api.editTransform.mockResolvedValue(slideResult)
    const apply = vi.fn()
    await flushTransforms(0, [{ sourceId: 'a', box }], apply)
    expect(api.editTransform).toHaveBeenCalledTimes(1)
    expect(api.batchEditTransform).not.toHaveBeenCalled()
    expect(apply).toHaveBeenCalledWith(slideResult)
  })

  it('several elements → one batchEditTransform', async () => {
    api.batchEditTransform.mockResolvedValue(slideResult)
    const apply = vi.fn()
    await flushTransforms(
      0,
      [
        { sourceId: 'a', box },
        { sourceId: 'b', box, groupId: 'g' },
      ],
      apply,
    )
    expect(api.editTransform).not.toHaveBeenCalled()
    expect(api.batchEditTransform).toHaveBeenCalledTimes(1)
    expect(api.batchEditTransform.mock.calls[0]![0].items).toHaveLength(2)
    expect(api.batchEditTransform.mock.calls[0]![0].items[1].groupId).toBe('g')
    expect(apply).toHaveBeenCalledTimes(1)
  })

  it('a refused batch falls back to per-element calls inside one history group', async () => {
    api.batchEditTransform.mockResolvedValue(null)
    api.editTransform.mockResolvedValue(slideResult)
    await flushTransforms(
      0,
      [
        { sourceId: 'a', box },
        { sourceId: 'b', box },
      ],
      vi.fn(),
    )
    expect(api.editTransform).toHaveBeenCalledTimes(2)
    expect(api.beginHistoryBatch).toHaveBeenCalledTimes(1)
    expect(api.endHistoryBatch).toHaveBeenCalledWith({ snapshot: false })
  })

  it('withHistoryGroup closes the group even when the body throws', async () => {
    await expect(
      withHistoryGroup(async () => {
        throw new Error('boom')
      }),
    ).rejects.toThrow('boom')
    expect(api.endHistoryBatch).toHaveBeenCalledTimes(1)
  })
})
