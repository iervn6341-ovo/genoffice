/**
 * Format painter / Paste Format (⌘⇧V): only the parts of the copied format that exist are written.
 * An outline-only (or text-only) paste must never touch the target's fill — it used to follow the
 * stroke with `editFill('none')` just to "refresh the page", wiping the fill.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import {
  applyBrushToElement,
  pasteFormat,
  onFormatBrushClick,
} from '../src/renderer/clipboard-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import type { BrushFormat } from '../src/renderer/format-brush'
import { registerTextCommit } from '../src/renderer/active-text-edit'

vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))

const calls: Array<{ api: string; arg: Record<string, unknown> }> = []
const slide = (tag: string) => ({ tag })

function installApi() {
  const rec = (api: string) => async (arg: Record<string, unknown>) => {
    calls.push({ api, arg })
    return slide(api)
  }
  ;(globalThis as unknown as { window: unknown }).window = {
    slidesApi: {
      editFill: rec('editFill'),
      editStroke: rec('editStroke'),
      editText: rec('editText'),
      applyTxn: async ({ ops }: { ops: Array<Record<string, unknown>> }) => {
        for (const op of ops) calls.push({ api: op.op as string, arg: op })
        return { applied: true, slides: [null, null, slide('txn')] }
      },
      beginHistoryBatch: async () => true,
      endHistoryBatch: async () => undefined,
    },
  }
}

const shapeNode = (id: string) => ({
  type: 'shape',
  sourceId: id,
  box: { x: 0, y: 0, w: 10, h: 10, rotationDeg: 0 },
  fill: { kind: 'solid', color: '#FFCC00' },
  text: { lines: [] },
})
const pictureNode = (id: string) => ({
  type: 'picture',
  sourceId: id,
  box: { x: 0, y: 0, w: 10, h: 10, rotationDeg: 0 },
})

function makeCtx(brush: BrushFormat, nodes: unknown[], selected: string[]) {
  const applied: unknown[] = []
  const ctx = {
    brushFormat: brush,
    brushMode: null,
    selectedIds: selected,
    current: 2,
    slide: { nodes },
    applySlide: (_i: number, s: unknown) => applied.push(s),
    setStatus: vi.fn(),
    setBrushMode: vi.fn(),
    setBrushFormat: vi.fn(),
  } as unknown as ActionCtx
  return { ctx, applied }
}

describe('Paste Format writes only what was copied', () => {
  beforeEach(() => {
    calls.length = 0
    installApi()
  })

  it('outline-only format (source fill was a gradient) leaves the target fill alone', async () => {
    const brush: BrushFormat = { sourceType: 'shape', stroke: { color: '#112233', widthPt: 3 } }
    const { ctx, applied } = makeCtx(brush, [shapeNode('a')], ['a'])
    await pasteFormat(ctx)
    expect(calls.map((c) => c.api)).toEqual(['setStroke'])
    expect(calls.some((c) => c.api === 'setFill')).toBe(false)
    expect(applied).toHaveLength(1) // the refreshed page still reaches the view
  })

  it('picture target gets the outline and never a fill write', async () => {
    const brush: BrushFormat = {
      sourceType: 'shape',
      fillColor: '#FF0000',
      stroke: { color: '#000000', widthPt: 1 },
    }
    const { ctx } = makeCtx(brush, [pictureNode('p')], ['p'])
    await pasteFormat(ctx)
    expect(calls.map((c) => c.api)).toEqual(['setStroke'])
  })

  it('fill + outline format writes each exactly once, in order, and surfaces every result', async () => {
    const brush: BrushFormat = {
      sourceType: 'shape',
      fillColor: '#FF0000',
      stroke: { color: '#000000', widthPt: 1 },
    }
    const { ctx, applied } = makeCtx(brush, [shapeNode('a'), shapeNode('b')], ['a', 'b'])
    await pasteFormat(ctx)
    expect(calls.map((c) => `${c.api}:${(c.arg.target as { el: string }).el}`)).toEqual([
      'setFill:a',
      'setStroke:a',
      'setFill:b',
      'setStroke:b',
    ])
    expect(calls[0]!.arg.fill).toBe('#FF0000')
    expect(applied).toHaveLength(2)
  })

  it('an explicit "no fill" source does clear the target fill (that is a real copied value)', async () => {
    const brush: BrushFormat = { sourceType: 'shape', fillColor: 'none' }
    const { ctx } = makeCtx(brush, [shapeNode('a')], ['a'])
    await pasteFormat(ctx)
    expect(calls).toEqual([
      { api: 'setFill', arg: { op: 'setFill', target: { slide: 2, el: 'a' }, fill: 'none' } },
    ])
  })

  it('format painter click shares the same behavior (outline-only keeps the fill)', async () => {
    const brush: BrushFormat = { sourceType: 'shape', stroke: { color: '#112233', widthPt: 2 } }
    const { ctx } = makeCtx(brush, [shapeNode('a')], [])
    await applyBrushToElement(ctx, 'a')
    expect(calls.map((c) => c.api)).toEqual(['setStroke'])
  })

  it('finishes the active text edit before reading the source format from the native document', async () => {
    const source = shapeNode('source')
    let committed = false
    const unregister = registerTextCommit(async () => {
      committed = true
    })
    window.slidesApi.getRenderSlide = vi.fn(async () => {
      expect(committed).toBe(true)
      return { nodes: [source], scale: 1 } as never
    })
    const { ctx } = makeCtx({ sourceType: 'shape' }, [source], ['source'])
    try {
      await onFormatBrushClick(ctx)
      expect(ctx.setBrushFormat).toHaveBeenCalledWith(
        expect.objectContaining({ fillColor: '#FFCC00', stroke: null }),
      )
      expect(ctx.setBrushMode).toHaveBeenCalledWith('once')
    } finally {
      unregister()
    }
  })

  it('retains the brush after a rejected atomic update and does not publish partial formatting', async () => {
    const { ctx, applied } = makeCtx(
      { sourceType: 'shape', fillColor: '#112233' },
      [shapeNode('a')],
      [],
    )
    ctx.brushMode = 'once'
    window.slidesApi.applyTxn = vi.fn(async () => ({ applied: false }))
    await applyBrushToElement(ctx, 'a')
    expect(applied).toHaveLength(0)
    expect(ctx.setBrushMode).not.toHaveBeenCalled()
    expect(ctx.setStatus).toHaveBeenCalled()
  })
})
