import { EventEmitter } from 'node:events'
import type { BrowserWindow } from 'electron'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { installNativeFrameEvents } from '../src/main/native-frame-events'

afterEach(() => vi.useRealTimers())

describe('native frame interactions', () => {
  it('dismisses once per resize/move burst and leaves native events untouched', () => {
    vi.useFakeTimers()
    const win = new EventEmitter()
    const dismiss = vi.fn()
    const event = { preventDefault: vi.fn() }
    installNativeFrameEvents(win as BrowserWindow, dismiss)
    for (let i = 0; i < 30; i++) {
      win.emit(i % 2 ? 'will-move' : 'will-resize', event)
      vi.advanceTimersByTime(16)
    }
    expect(dismiss).toHaveBeenCalledTimes(1)
    expect(event.preventDefault).not.toHaveBeenCalled()
    expect(vi.getTimerCount()).toBe(1)
    vi.advanceTimersByTime(150)
    expect(vi.getTimerCount()).toBe(0)
    win.emit('will-move', event)
    expect(dismiss).toHaveBeenCalledTimes(2)
  })

  it('keeps separate windows independent and removes pending work on close', () => {
    vi.useFakeTimers()
    const a = new EventEmitter()
    const b = new EventEmitter()
    const dismissA = vi.fn()
    const dismissB = vi.fn()
    installNativeFrameEvents(a as BrowserWindow, dismissA)
    installNativeFrameEvents(b as BrowserWindow, dismissB)
    a.emit('will-resize', {})
    expect(dismissA).toHaveBeenCalledTimes(1)
    expect(dismissB).not.toHaveBeenCalled()
    a.emit('closed')
    expect(vi.getTimerCount()).toBe(0)
    expect(a.listenerCount('will-move')).toBe(0)
    expect(a.listenerCount('will-resize')).toBe(0)
    b.emit('will-move', {})
    expect(dismissB).toHaveBeenCalledTimes(1)
    b.emit('closed')
    expect(vi.getTimerCount()).toBe(0)
  })
})
