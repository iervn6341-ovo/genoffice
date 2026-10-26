// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'

import { installKeepGridFocus, shouldKeepGridFocus } from '../src/renderer/keep-grid-focus'

function build() {
  document.body.innerHTML = `
    <header data-keep-edit="">
      <button id="bold">B</button>
      <button id="icon"><svg id="glyph"></svg></button>
      <input id="size" />
      <label id="lbl"><input id="color" type="color" /></label>
      <select id="sel"><option>a</option></select>
    </header>
    <main><button id="outside">x</button><div id="grid" tabindex="0"></div></main>`
  const el = (id: string) => document.getElementById(id)!
  return el
}

function press(target: Element): MouseEvent {
  const e = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })
  target.dispatchEvent(e)
  return e
}

afterEach(() => {
  document.body.innerHTML = ''
  vi.useRealTimers()
})

describe('shouldKeepGridFocus', () => {
  it('keeps focus for plain ribbon controls only', () => {
    const el = build()
    expect(shouldKeepGridFocus(el('bold'))).toBe(true)
    expect(shouldKeepGridFocus(el('glyph'))).toBe(true)
    expect(shouldKeepGridFocus(el('size'))).toBe(false)
    expect(shouldKeepGridFocus(el('sel'))).toBe(false)
    expect(shouldKeepGridFocus(el('lbl'))).toBe(false)
    expect(shouldKeepGridFocus(el('outside'))).toBe(false)
    expect(shouldKeepGridFocus(null)).toBe(false)
  })
})

describe('installKeepGridFocus', () => {
  it('a ribbon button press does not move focus off the grid', () => {
    const el = build()
    const off = installKeepGridFocus({ focusGrid: () => {} })
    el('grid').focus()
    expect(press(el('bold')).defaultPrevented).toBe(true)
    expect(press(el('glyph')).defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(el('grid'))
    off()
  })

  it('fields, labels and presses outside the ribbon keep their normal focus behaviour', () => {
    const el = build()
    const off = installKeepGridFocus({ focusGrid: () => {} })
    expect(press(el('size')).defaultPrevented).toBe(false)
    expect(press(el('lbl')).defaultPrevented).toBe(false)
    expect(press(el('outside')).defaultPrevented).toBe(false)
    off()
  })

  it('only the primary button is guarded', () => {
    const el = build()
    const off = installKeepGridFocus({ focusGrid: () => {} })
    const right = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 2 })
    el('bold').dispatchEvent(right)
    expect(right.defaultPrevented).toBe(false)
    off()
  })

  it('a pending value in a ribbon box commits on blur, then the grid gets focus back', () => {
    vi.useFakeTimers()
    const el = build()
    const focusGrid = vi.fn()
    const off = installKeepGridFocus({ focusGrid })
    el('size').focus()
    expect(press(el('bold')).defaultPrevented).toBe(false)
    expect(focusGrid).not.toHaveBeenCalled()
    document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }))
    vi.runAllTimers()
    expect(focusGrid).toHaveBeenCalledTimes(1)
    off()
  })

  it('uninstalling stops guarding', () => {
    const el = build()
    const off = installKeepGridFocus({ focusGrid: () => {} })
    off()
    expect(press(el('bold')).defaultPrevented).toBe(false)
  })
})
