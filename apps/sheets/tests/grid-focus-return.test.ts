// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { scheduleGridFocusReturn, setGridFocusHandler } from '../src/renderer/grid-focus-return'

beforeEach(() => vi.useFakeTimers())
afterEach(() => {
  vi.useRealTimers()
  setGridFocusHandler(null)
  document.body.innerHTML = ''
})

describe('grid focus return after an object is deleted', () => {
  it('hands focus to the grid when it fell to <body>', () => {
    const focus = vi.fn()
    setGridFocusHandler(focus)
    scheduleGridFocusReturn()
    vi.runAllTimers()
    expect(focus).toHaveBeenCalledTimes(1)
  })

  it('never takes focus from a field the user is typing in', () => {
    const focus = vi.fn()
    setGridFocusHandler(focus)
    document.body.innerHTML = '<input id="chat" />'
    document.getElementById('chat')!.focus()
    scheduleGridFocusReturn()
    vi.runAllTimers()
    expect(focus).not.toHaveBeenCalled()
  })

  it('is a no-op without a registered handler', () => {
    scheduleGridFocusReturn()
    expect(() => vi.runAllTimers()).not.toThrow()
  })
})

describe('delayed loss of focus', () => {
  it('still hands focus back when the node only disappears a moment later', () => {
    const focus = vi.fn()
    setGridFocusHandler(focus)
    document.body.innerHTML = '<div id="obj" tabindex="0"></div>'
    document.getElementById('obj')!.focus()
    scheduleGridFocusReturn()
    vi.advanceTimersByTime(10)
    expect(focus).not.toHaveBeenCalled()
    document.getElementById('obj')!.remove() // the unmount lands
    vi.advanceTimersByTime(200)
    expect(focus).toHaveBeenCalledTimes(1)
    vi.runAllTimers()
    expect(focus).toHaveBeenCalledTimes(1)
  })
})
