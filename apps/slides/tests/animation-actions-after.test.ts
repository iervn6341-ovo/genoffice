import { afterEach, expect, it, vi } from 'vitest'
import { patchAnimation } from '../src/renderer/animation-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import type { AnimationItem } from '../src/shared/ipc'

const effect = (nativeId: number): AnimationItem => ({
  sourceId: 'card',
  targetName: 'Card',
  nativeId,
  effect: 'appear',
  trigger: 'onClick',
  durationMs: 0,
  delayMs: 0,
})
function setup(initial: AnimationItem[], serverInitial: AnimationItem[]) {
  let server = structuredClone(serverInitial)
  const ctx = {
    current: 8,
    path: '/test.pptx',
    animations: initial,
    setAnimations: vi.fn((items: AnimationItem[]) => {
      ctx.animations = items
    }),
    setDirty: vi.fn(),
    setStatus: vi.fn(),
    setSelAnim: vi.fn(),
  } as unknown as ActionCtx
  const set = vi.fn(async ({ items }: { items: AnimationItem[] }) => {
    if (items.some((a, i) => a.nativeId != null && a.nativeId !== server[i]?.nativeId)) return false
    server = items.map((a, i) => ({ ...a, nativeId: 100 + i, targetName: 'Card' }))
    return true
  })
  Object.assign(window, {
    slidesApi: {
      getAnimations: vi.fn(async () => structuredClone(server)),
      setAnimations: set,
      applyTxn: vi.fn(async () => ({
        applied: false,
        failures: [{ index: 0, error: 'stale native effect' }],
      })),
    },
  })
  return { ctx, set, read: () => server }
}
afterEach(() => vi.restoreAllMocks())

it('refreshes an outdated native effect before writing the after-animation selection', async () => {
  const { ctx, set, read } = setup([effect(5)], [effect(50)])
  await patchAnimation(ctx, 0, { after: 'hideAfter' })
  expect(set.mock.calls[0]![0].items[0]!.nativeId).toBe(50)
  expect(read()[0]!.after).toBe('hideAfter')
  expect(ctx.animations[0]!.after).toBe('hideAfter')
})

it('applies consecutive choices in order using the current native IDs each time', async () => {
  const { ctx, read } = setup([effect(5)], [effect(5)])
  const first = patchAnimation(ctx, 0, { after: 'hideAfter' })
  const second = patchAnimation(ctx, 0, { after: 'hideOnNextClick' })
  await Promise.all([first, second])
  expect(read()[0]!.after).toBe('hideOnNextClick')
  expect(ctx.animations[0]!.after).toBe('hideOnNextClick')
})

it('resolves the same native effect after save has refreshed its source ID', async () => {
  const original = { ...effect(5), sourceId: 'old-source' }
  const fresh = { ...effect(5), sourceId: 'saved-source' }
  const { ctx, set, read } = setup([original], [fresh])
  await patchAnimation(ctx, 0, { after: 'hideOnNextClick' })
  expect(set.mock.calls[0]![0].items[0]!.sourceId).toBe('saved-source')
  expect(read()[0]!.after).toBe('hideOnNextClick')
})

it('preserves another stacked effect when native IDs have both changed', async () => {
  const { ctx, set, read } = setup([effect(5), effect(6)], [effect(50), effect(60)])
  await patchAnimation(ctx, 1, { after: 'hideAfter' })
  expect(set.mock.calls[0]![0].items.map((a) => a.nativeId)).toEqual([50, 60])
  expect(read()[0]!.after).toBeUndefined()
  expect(read()[1]!.after).toBe('hideAfter')
})

it('rolls back the visible choice and reports a rejected update instead of silently resetting it', async () => {
  const { ctx, set } = setup([effect(5)], [effect(5)])
  set.mockResolvedValue(false)
  await patchAnimation(ctx, 0, { after: 'hideAfter' })
  expect(ctx.animations[0]!.after).toBeUndefined()
  expect(ctx.setDirty).not.toHaveBeenCalled()
  expect(ctx.setStatus).toHaveBeenCalledWith(expect.stringContaining('stale native effect'))
})

it('does not apply a removed effect selection to the remaining effect on the same shape', async () => {
  const { ctx, set } = setup([effect(5), effect(6)], [effect(6)])
  await patchAnimation(ctx, 0, { after: 'hideAfter' })
  expect(set).not.toHaveBeenCalled()
  expect(ctx.animations).toHaveLength(1)
  expect(ctx.animations[0]!.after).toBeUndefined()
})

it('does not publish an earlier page update into the page selected while its write was pending', async () => {
  const { ctx, set } = setup([effect(5)], [effect(5)])
  let release!: () => void
  let started!: () => void
  const began = new Promise<void>((resolve) => {
    started = resolve
  })
  const gate = new Promise<void>((resolve) => {
    release = resolve
  })
  const actualWrite = set.getMockImplementation()!
  set.mockImplementationOnce(async (args) => {
    started()
    await gate
    return actualWrite(args)
  })
  let active = ctx
  ctx.getLatestContext = () => active
  const work = patchAnimation(ctx, 0, { after: 'hideAfter' })
  await began
  const otherPageUpdates = vi.fn()
  active = { ...ctx, current: 9, setAnimations: otherPageUpdates }
  release()
  await work
  expect(otherPageUpdates).not.toHaveBeenCalled()
})

it('does not write into a different document opened while the latest list was being requested', async () => {
  const { ctx, set } = setup([effect(5)], [effect(5)])
  let release!: (value: AnimationItem[]) => void
  let started!: () => void
  const began = new Promise<void>((resolve) => {
    started = resolve
  })
  window.slidesApi.getAnimations = vi.fn(() => {
    started()
    return new Promise<AnimationItem[]>((resolve) => {
      release = resolve
    })
  })
  let active = ctx
  ctx.getLatestContext = () => active
  const work = patchAnimation(ctx, 0, { after: 'hideAfter' })
  await began
  active = { ...ctx, path: '/other.pptx' }
  release([effect(5)])
  await work
  expect(set).not.toHaveBeenCalled()
})
