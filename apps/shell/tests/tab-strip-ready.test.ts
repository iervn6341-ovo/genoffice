import { afterEach, describe, expect, it, vi } from 'vitest'
import { TabStripReady } from '../src/main/tab-strip-ready'

afterEach(() => vi.useRealTimers())

describe('detached title-bar readiness', () => {
  it('waits for the transferred tab, not an earlier Home paint', async () => {
    const ready = new TabStripReady()
    ready.acknowledge('home')
    expect(await ready.wait('home')).toBe(true)
    let painted = false
    const pending = ready.wait('3-t1').then((value) => {
      painted = value
    })
    await Promise.resolve()
    expect(painted).toBe(false)
    ready.acknowledge('home')
    await Promise.resolve()
    expect(painted).toBe(false)
    ready.acknowledge('3-t1')
    await pending
    expect(painted).toBe(true)
  })

  it('bounds failed-render waits and releases them when a window closes', async () => {
    vi.useFakeTimers()
    const ready = new TabStripReady()
    const expired = ready.wait('home')
    vi.advanceTimersByTime(8000)
    expect(await expired).toBe(false)
    const pending = ready.wait('3-t1')
    ready.dispose()
    expect(await pending).toBe(false)
    expect(await ready.wait('3-t1')).toBe(false)
    expect(vi.getTimerCount()).toBe(0)
  })
})
