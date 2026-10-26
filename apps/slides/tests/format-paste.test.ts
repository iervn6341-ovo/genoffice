/**
 * Format painter / Paste Format (⌘⇧V): only the parts of the copied format that exist are written.
 * An outline-only (or text-only) paste must never touch the target's fill — it used to follow the
 * stroke with `editFill('none')` just to "refresh the page", wiping the fill.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { applyBrushToElement, pasteFormat } from '../src/renderer/clipboard-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import type { BrushFormat } from '../src/renderer/format-brush'

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
    expect(calls.map((c) => c.api)).toEqual(['editStroke'])
    expect(calls.some((c) => c.api === 'editFill')).toBe(false)
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
    expect(calls.map((c) => c.api)).toEqual(['editStroke'])
  })

  it('fill + outline format writes each exactly once, in order, and surfaces every result', async () => {
    const brush: BrushFormat = {
      sourceType: 'shape',
      fillColor: '#FF0000',
      stroke: { color: '#000000', widthPt: 1 },
    }
    const { ctx, applied } = makeCtx(brush, [shapeNode('a'), shapeNode('b')], ['a', 'b'])
    await pasteFormat(ctx)
    expect(calls.map((c) => `${c.api}:${c.arg.sourceId as string}`)).toEqual([
      'editFill:a',
      'editStroke:a',
      'editFill:b',
      'editStroke:b',
    ])
    expect(calls[0]!.arg.fill).toBe('#FF0000')
    expect(applied).toHaveLength(4)
  })

  it('an explicit "no fill" source does clear the target fill (that is a real copied value)', async () => {
    const brush: BrushFormat = { sourceType: 'shape', fillColor: 'none' }
    const { ctx } = makeCtx(brush, [shapeNode('a')], ['a'])
    await pasteFormat(ctx)
    expect(calls).toEqual([
      { api: 'editFill', arg: { slideIndex: 2, sourceId: 'a', fill: 'none' } },
    ])
  })

  it('format painter click shares the same behavior (outline-only keeps the fill)', async () => {
    const brush: BrushFormat = { sourceType: 'shape', stroke: { color: '#112233', widthPt: 2 } }
    const { ctx } = makeCtx(brush, [shapeNode('a')], [])
    await applyBrushToElement(ctx, 'a')
    expect(calls.map((c) => c.api)).toEqual(['editStroke'])
  })
})
