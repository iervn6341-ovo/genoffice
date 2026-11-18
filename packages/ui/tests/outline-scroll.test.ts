import { afterEach, describe, expect, it, vi } from 'vitest'
import { activeOutlineId, observeOutlineScroll } from '../src/outline-scroll'

afterEach(() => {
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

describe('outline reading position', () => {
  const anchors = [
    { id: 'parent', top: 100 },
    { id: 'child', top: 100 },
    { id: 'later', top: 800 },
  ]
  it('follows the viewport in both directions and keeps the parent for identical destinations', () => {
    expect(activeOutlineId(anchors, 0)).toBe('parent')
    expect(activeOutlineId(anchors, 799)).toBe('parent')
    expect(activeOutlineId(anchors, 800)).toBe('later')
    expect(activeOutlineId(anchors, 500)).toBe('parent')
    expect(activeOutlineId([], 0)).toBeNull()
  })

  it('coalesces scroll bursts without remeasuring; reflows refresh once and cleanup cancels callbacks', () => {
    vi.useFakeTimers()
    const frames = new Map<number, FrameRequestCallback>()
    let sequence = 0
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => {
      frames.set(++sequence, cb)
      return sequence
    })
    vi.stubGlobal('cancelAnimationFrame', (id: number) => frames.delete(id))
    const flush = () => {
      const batch = [...frames.values()]
      frames.clear()
      batch.forEach((cb) => cb(0))
    }
    let resized = () => {}
    const disconnect = vi.fn()
    vi.stubGlobal(
      'ResizeObserver',
      class {
        constructor(cb: () => void) {
          resized = cb
        }
        observe() {}
        disconnect = disconnect
      },
    )
    vi.stubGlobal('window', { setTimeout })
    vi.stubGlobal('document', { fonts: new EventTarget() })
    const element = Object.assign(new EventTarget(), {
      scrollTop: 0,
      clientHeight: 500,
      scrollHeight: 1600,
    }) as unknown as HTMLElement
    let positions = [...anchors]
    const measure = vi.fn(() => positions)
    const changed = vi.fn()
    const stop = observeOutlineScroll(element, null, measure, changed)
    flush()
    expect(changed).toHaveBeenLastCalledWith('parent')
    element.scrollTop = 810
    for (let i = 0; i < 100; i++) element.dispatchEvent(new Event('scroll'))
    expect(frames.size).toBe(1)
    flush()
    expect(changed).toHaveBeenLastCalledWith('later')
    expect(measure).toHaveBeenCalledTimes(1)
    positions = [
      { id: 'moved', top: 0 },
      { id: 'later', top: 1200 },
    ]
    for (let i = 0; i < 100; i++) resized()
    vi.advanceTimersByTime(80)
    flush()
    expect(measure).toHaveBeenCalledTimes(2)
    expect(changed).toHaveBeenLastCalledWith('moved')
    element.scrollTop = 1100
    element.dispatchEvent(new Event('scroll'))
    flush()
    expect(changed).toHaveBeenLastCalledWith('later')
    element.dispatchEvent(new Event('scroll'))
    resized()
    const count = changed.mock.calls.length
    stop()
    vi.runAllTimers()
    flush()
    element.dispatchEvent(new Event('scroll'))
    expect(changed).toHaveBeenCalledTimes(count)
    expect(disconnect).toHaveBeenCalledOnce()
    expect(frames.size).toBe(0)
  })
})
