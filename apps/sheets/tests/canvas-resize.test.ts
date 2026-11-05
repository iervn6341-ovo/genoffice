import { afterEach, describe, expect, it, vi } from 'vitest'
import { Subject } from 'rxjs'
import { UniverInstanceType } from '@univerjs/core'
import type { IRenderManagerService } from '@univerjs/engine-render'
import { coalesceCanvasResize, installCanvasResizeCoalescing } from '../src/renderer/canvas-resize'

afterEach(() => vi.useRealTimers())

describe('canvas resize coalescing', () => {
  it('mounts immediately, limits a continuous burst, and paints the final size', () => {
    vi.useFakeTimers()
    let size = 100
    const painted: number[] = []
    const engine = {
      resize() {
        painted.push(size)
      },
    }
    const dispose = coalesceCanvasResize(engine)
    engine.resize()
    expect(painted).toEqual([100])
    for (let i = 1; i <= 10; i++) {
      size = 100 + i
      engine.resize()
      vi.advanceTimersByTime(16)
    }
    vi.runAllTimers()
    expect(painted).toEqual([100, 105, 110])
    dispose()
  })

  it('cancels pending work and restores immediate resizing on disposal', () => {
    vi.useFakeTimers()
    const resize = vi.fn()
    const engine = { resize }
    const dispose = coalesceCanvasResize(engine)
    engine.resize()
    engine.resize()
    dispose()
    vi.runAllTimers()
    expect(resize).toHaveBeenCalledTimes(1)
    expect(engine.resize).toBe(resize)
    engine.resize()
    expect(resize).toHaveBeenCalledTimes(2)
  })

  it('follows replacement workbook engines, leaves text editors immediate, and cancels old work', () => {
    vi.useFakeTimers()
    const created$ = new Subject<any>()
    const disposed$ = new Subject<string>()
    const sheet = {
      unitId: 'old',
      type: UniverInstanceType.UNIVER_SHEET,
      engine: { resize: vi.fn() },
    }
    const editor = {
      unitId: 'editor',
      type: UniverInstanceType.UNIVER_DOC,
      engine: { resize: vi.fn() },
    }
    const oldResize = sheet.engine.resize
    const editorResize = editor.engine.resize
    const manager = {
      getRenderAll: () =>
        new Map([
          ['old', sheet],
          ['editor', editor],
        ]),
      created$,
      disposed$,
    }
    const dispose = installCanvasResizeCoalescing(manager as unknown as IRenderManagerService)
    expect(editor.engine.resize).toBe(editorResize)
    sheet.engine.resize()
    sheet.engine.resize()
    disposed$.next('old')
    vi.runAllTimers()
    expect(oldResize).toHaveBeenCalledTimes(1)
    expect(sheet.engine.resize).toBe(oldResize)
    const replacement = {
      unitId: 'new',
      type: UniverInstanceType.UNIVER_SHEET,
      engine: { resize: vi.fn() },
    }
    const newResize = replacement.engine.resize
    created$.next(replacement)
    replacement.engine.resize()
    replacement.engine.resize()
    vi.advanceTimersByTime(80)
    expect(newResize).toHaveBeenCalledTimes(2)
    dispose()
    expect(replacement.engine.resize).toBe(newResize)
    expect(created$.observed).toBe(false)
    expect(disposed$.observed).toBe(false)
  })
})
