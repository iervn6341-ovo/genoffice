import { describe, expect, it, vi } from 'vitest'

// file-actions pulls export-render → pptx-render → konva → native `canvas`
// (missing in CI). Mock the UI-only modules so this unit test exercises
// only the save-serialization queue.
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))
vi.mock('../src/renderer/components/toast-bus', () => ({ showToast: vi.fn() }))
vi.mock('../src/renderer/i18n/locale', () => ({ t: (k: string) => k }))

import { save, saveAs } from '../src/renderer/file-actions'
import { patchAnimation } from '../src/renderer/animation-actions'
import { createTransformBatcher } from '../src/renderer/history-group'
import { queueFormatEdit } from '../src/renderer/format-edit-queue'
import type { ActionCtx } from '../src/renderer/action-context'
import type { AnimationItem } from '../src/shared/ipc'

function ctx(): ActionCtx {
  return {
    flushActiveEditRef: { current: () => Promise.resolve() },
    editingActiveRef: { current: false },
    flushNotes: () => Promise.resolve(),
    slides: [],
    current: 0,
    setSlides: () => {},
    setSelectedIds: () => {},
    setEnteredGroupId: () => {},
    setEditing: () => {},
    setEditingCell: () => {},
    setDirty: () => {},
    setPath: () => {},
    setStatus: () => {},
    path: '/test/deck.pptx',
  } as unknown as ActionCtx
}

describe('slides save serialization', () => {
  it.each([
    ['move', 'geometry'],
    ['format brush', 'format'],
  ] as const)('waits for an unfinished %s before saving', async (_name, kind) => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => {
      release = resolve
    })
    let committed = false
    const action = async () => {
      await gate
      committed = true
    }
    const edit =
      kind === 'geometry'
        ? createTransformBatcher(action).push(0, {
            sourceId: 'node',
            box: { x: 10, y: 20, w: 30, h: 40, rotationDeg: 25 },
          })
        : queueFormatEdit(action)
    const disk = vi.fn(async () => {
      expect(committed).toBe(true)
      return { ok: true }
    })
    vi.stubGlobal('window', { slidesApi: { save: disk } })
    try {
      const saving = save(ctx)
      await Promise.resolve()
      expect(disk).not.toHaveBeenCalled()
      release()
      await edit
      expect(await saving).toBe(true)
    } finally {
      release()
      vi.unstubAllGlobals()
    }
  })
  it.each([
    ['Save', save],
    ['Save As', saveAs],
  ] as const)(
    '%s waits for the after-animation choice before writing the document',
    async (_name, saveDocument) => {
      let stored: AnimationItem[] = [
        {
          sourceId: 'card',
          targetName: 'Card',
          effect: 'appear',
          trigger: 'onClick',
          durationMs: 0,
          delayMs: 0,
          nativeId: 5,
        },
      ]
      const current = ctx()
      current.animations = stored
      current.setAnimations = (items) => {
        current.animations = items as AnimationItem[]
      }
      let release!: () => void
      let started!: () => void
      const began = new Promise<void>((resolve) => {
        started = resolve
      })
      const gate = new Promise<void>((resolve) => {
        release = resolve
      })
      const diskWrite = vi.fn(async () => {
        expect(stored[0]!.after).toBe('hideOnNextClick')
        return { ok: true }
      })
      vi.stubGlobal('window', {
        slidesApi: {
          getAnimations: async () => structuredClone(stored),
          setAnimations: async ({ items }: { items: AnimationItem[] }) => {
            started()
            await gate
            stored = items
            return true
          },
          save: diskWrite,
          saveAs: diskWrite,
        },
      })
      try {
        const edit = patchAnimation(current, 0, { after: 'hideOnNextClick' })
        const saving = saveDocument(() => current)
        await began
        expect(diskWrite).not.toHaveBeenCalled()
        release()
        await edit
        expect(await saving).toBe(true)
        expect(diskWrite).toHaveBeenCalledTimes(1)
      } finally {
        release()
        vi.unstubAllGlobals()
      }
    },
  )

  it('queues concurrent saves so only one IPC write runs at a time', async () => {
    let inFlight = 0
    let maxConcurrent = 0
    const mockSave = vi.fn(async () => {
      inFlight += 1
      maxConcurrent = Math.max(maxConcurrent, inFlight)
      // Simulate a slow write.
      await new Promise((resolve) => setTimeout(resolve, 50))
      inFlight -= 1
      return { ok: true }
    })
    vi.stubGlobal('window', { slidesApi: { save: mockSave } })

    // Fire three saves simultaneously — they must queue, not overlap.
    const [a, b, c] = await Promise.all([save(ctx), save(ctx, true), save(ctx)])

    expect(a).toBe(true)
    expect(b).toBe(true)
    expect(c).toBe(true)
    expect(mockSave).toHaveBeenCalledTimes(3)
    expect(maxConcurrent).toBe(1)

    vi.unstubAllGlobals()
  })

  it('serializes save against saveAs so their IPC writes never overlap', async () => {
    let inFlight = 0
    let maxConcurrent = 0
    const track = async () => {
      inFlight += 1
      maxConcurrent = Math.max(maxConcurrent, inFlight)
      // Simulate a slow write.
      await new Promise((resolve) => setTimeout(resolve, 50))
      inFlight -= 1
      return { ok: true }
    }
    const mockSave = vi.fn(track)
    const mockSaveAs = vi.fn(track)
    vi.stubGlobal('window', { slidesApi: { save: mockSave, saveAs: mockSaveAs } })

    // A manual save racing Save As (or close-save) must queue behind it.
    const [saved, _] = await Promise.all([save(ctx, true), saveAs(ctx)])

    expect(saved).toBe(true)
    expect(mockSave).toHaveBeenCalledTimes(1)
    expect(mockSaveAs).toHaveBeenCalledTimes(1)
    expect(maxConcurrent).toBe(1)

    vi.unstubAllGlobals()
  })

  it('reports failure without blocking the queue for the next caller', async () => {
    let calls = 0
    const mockSave = vi.fn(async () => {
      calls += 1
      if (calls === 1) return { ok: false, error: 'disk full' }
      return { ok: true }
    })
    vi.stubGlobal('window', { slidesApi: { save: mockSave } })

    const first = save(ctx, true)
    const second = save(ctx, true)
    const [r1, r2] = await Promise.all([first, second])

    expect(r1).toBe(false)
    expect(r2).toBe(true)
    expect(mockSave).toHaveBeenCalledTimes(2)

    vi.unstubAllGlobals()
  })

  it('resolves the context when a queued pass runs, not when it was queued', async () => {
    let saves = 0
    const mockSave = vi.fn(async () => {
      await new Promise((resolve) => setTimeout(resolve, 20))
      saves += 1
      return { ok: true }
    })
    vi.stubGlobal('window', { slidesApi: { save: mockSave } })

    const seen: number[] = []
    const getCtx = () => {
      seen.push(saves)
      return ctx()
    }
    await Promise.all([save(getCtx, true), save(getCtx, true)])

    expect(seen).toEqual([0, 1])

    vi.unstubAllGlobals()
  })
})
