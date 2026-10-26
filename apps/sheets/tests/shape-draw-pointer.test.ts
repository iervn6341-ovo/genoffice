// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'

import { startSheetShapeDraw } from '../src/renderer/shape-draw'
import type { UniverRuntime } from '../src/renderer/univer-state'

afterEach(() => {
  window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
  document.body.innerHTML = ''
})

describe('shape draw mode and the grid canvas', () => {
  it('the drawing press never reaches the grid, so it cannot start a cell selection', () => {
    document.body.innerHTML = '<div id="univer-container"><canvas id="grid"></canvas></div>'
    const grid = document.getElementById('grid')!
    const gridPointerDown = vi.fn()
    grid.addEventListener('pointerdown', gridPointerDown)

    startSheetShapeDraw({} as UniverRuntime, 'rect', () => {})
    const press = new MouseEvent('pointerdown', { bubbles: true, cancelable: true, button: 0 })
    grid.dispatchEvent(press)
    expect(gridPointerDown).not.toHaveBeenCalled()
    // compat mouse events are not cancelled by it (the draw gesture itself listens for mousedown)
    expect(press.defaultPrevented).toBe(false)

    // once the mode ends (Esc) the grid gets its presses again
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    grid.dispatchEvent(
      new MouseEvent('pointerdown', { bubbles: true, cancelable: true, button: 0 }),
    )
    expect(gridPointerDown).toHaveBeenCalledTimes(1)
  })
})
