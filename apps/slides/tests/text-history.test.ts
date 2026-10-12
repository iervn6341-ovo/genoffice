// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import {
  createTextHistory,
  redoTextEdit,
  setActiveTextHistory,
  undoTextEdit,
} from '../src/renderer/text-history'

/** the overlay's markup in miniature: paragraph divs that carry level / margin as attributes */
function editor(): HTMLElement {
  const root = document.createElement('div')
  root.contentEditable = 'true'
  root.innerHTML = '<div data-level="1" style="margin-left:128px">one</div><div>two</div>'
  document.body.append(root)
  return root
}

const state = (root: HTMLElement) =>
  [...root.children].map((c) => `${c.textContent}|${(c as HTMLElement).dataset.level ?? '-'}`)

beforeEach(() => vi.useFakeTimers())
afterEach(() => {
  vi.useRealTimers()
  document.body.innerHTML = ''
  setActiveTextHistory(null)
})

describe('text overlay undo / redo', () => {
  it('takes back an attribute edit the browser would never remember', async () => {
    const root = editor()
    const history = createTextHistory(root)
    const first = root.children[0] as HTMLElement
    delete first.dataset.level // what Shift+Tab does
    first.style.marginLeft = '48px'
    await vi.advanceTimersByTimeAsync(600)
    expect(state(root)).toEqual(['one|-', 'two|-'])
    history.undo()
    expect(state(root)).toEqual(['one|1', 'two|-'])
    expect((root.children[0] as HTMLElement).style.marginLeft).toBe('128px')
    history.redo()
    expect(state(root)).toEqual(['one|-', 'two|-'])
    history.dispose()
  })

  it('typing in one burst is one step, a flush starts the next one', async () => {
    const root = editor()
    const history = createTextHistory(root)
    const text = root.children[1]!.firstChild as Text
    text.data = 'two a'
    text.data = 'two ab'
    history.flush() // a discrete key (Enter, Tab...) closes the burst
    const added = document.createElement('div')
    added.textContent = 'three'
    root.append(added)
    await vi.advanceTimersByTimeAsync(600)
    expect(root.children.length).toBe(3)
    history.undo()
    expect(root.children.length).toBe(2)
    expect(root.children[1]!.textContent).toBe('two ab')
    history.undo()
    expect(root.children[1]!.textContent).toBe('two')
    history.undo() // nothing older: stays put
    expect(root.children[1]!.textContent).toBe('two')
    history.dispose()
  })

  it('a new edit after an undo drops the redo tail', async () => {
    const root = editor()
    const history = createTextHistory(root)
    ;(root.children[1]!.firstChild as Text).data = 'changed'
    await vi.advanceTimersByTimeAsync(600)
    history.undo()
    ;(root.children[1]!.firstChild as Text).data = 'other'
    await vi.advanceTimersByTimeAsync(600)
    history.redo()
    expect(root.children[1]!.textContent).toBe('other')
    history.dispose()
  })

  it('the module-level undo / redo reach the open editor, and report none when closed', async () => {
    expect(undoTextEdit()).toBe(false)
    const root = editor()
    const history = createTextHistory(root)
    setActiveTextHistory(history)
    ;(root.children[1]!.firstChild as Text).data = 'x'
    await vi.advanceTimersByTimeAsync(600)
    expect(undoTextEdit()).toBe(true)
    expect(root.children[1]!.textContent).toBe('two')
    expect(redoTextEdit()).toBe(true)
    expect(root.children[1]!.textContent).toBe('x')
    history.dispose()
  })
})
