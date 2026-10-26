// @vitest-environment jsdom
// The shared keep-edit behavior (ribbon / pane use must not end an in-place text edit or drop its
// selection) — used by the Slides text editor and meant for any contentEditable editor.
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  classifyEditPress,
  createSelectionKeeper,
  handleEditorBlur,
  installKeepEditGuard,
} from '../src/keep-edit'

function editor(text: string): HTMLDivElement {
  const frame = document.createElement('div')
  const div = document.createElement('div')
  div.contentEditable = 'true'
  div.tabIndex = 0
  // jsdom does not derive isContentEditable from the attribute
  Object.defineProperty(div, 'isContentEditable', { value: true })
  div.textContent = text
  frame.appendChild(div)
  document.body.appendChild(frame)
  return div
}

function select(div: HTMLElement, a: number, b: number): void {
  const r = document.createRange()
  r.setStart(div.firstChild!, a)
  r.setEnd(div.firstChild!, b)
  window.getSelection()!.removeAllRanges()
  window.getSelection()!.addRange(r)
}

const zoneWith = (...kids: HTMLElement[]) => {
  const zone = document.createElement('div')
  zone.setAttribute('data-keep-edit', '')
  zone.append(...kids)
  document.body.appendChild(zone)
  return zone
}

afterEach(() => {
  document.body.innerHTML = ''
})

describe('classifyEditPress', () => {
  it('editor frame / zone button / zone field / outside', () => {
    const div = editor('x')
    const btn = document.createElement('button')
    const input = document.createElement('input')
    zoneWith(btn, input)
    const other = document.createElement('div')
    document.body.appendChild(other)
    expect(classifyEditPress(div.parentElement, div)).toBe('in-editor')
    expect(classifyEditPress(div.parentElement, btn)).toBe('neutral')
    expect(classifyEditPress(div.parentElement, input)).toBe('neutral-field')
    expect(classifyEditPress(div.parentElement, other)).toBe('outside')
  })
})

describe('selection keeper', () => {
  it('remember → restore re-focuses the editor with the snapshot; forget drops it', () => {
    const div = editor('AAABBB')
    div.focus()
    select(div, 1, 4)
    const keeper = createSelectionKeeper('test-kept')
    keeper.remember(div)
    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    window.getSelection()!.removeAllRanges()
    expect(keeper.restore()).toBe(true)
    expect(document.activeElement).toBe(div)
    expect(window.getSelection()!.toString()).toBe('AAB')
    keeper.forget()
    expect(keeper.restore()).toBe(false)
  })

  it('ignores a selection that is not inside the editor', () => {
    const div = editor('inside')
    const other = document.createElement('p')
    other.textContent = 'outside'
    document.body.appendChild(other)
    const r = document.createRange()
    r.selectNodeContents(other)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(r)
    const keeper = createSelectionKeeper('test-kept')
    keeper.remember(div)
    expect(keeper.peek()).toBeNull()
  })
})

describe('handleEditorBlur', () => {
  it('absorbs a blur into a zone (and remembers), lets any other blur through', () => {
    const div = editor('hello')
    div.focus()
    select(div, 0, 5)
    const keeper = createSelectionKeeper('test-kept')
    const btn = document.createElement('button')
    zoneWith(btn)
    expect(handleEditorBlur({ relatedTarget: btn, currentTarget: div }, keeper)).toBe(true)
    expect(keeper.peek()?.range.toString()).toBe('hello')
    expect(handleEditorBlur({ relatedTarget: null, currentTarget: div }, keeper)).toBe(false)
    expect(handleEditorBlur({ relatedTarget: document.body, currentTarget: div }, keeper)).toBe(
      false,
    )
  })
})

describe('installKeepEditGuard', () => {
  const press = (el: Element) => {
    const ev = new MouseEvent('mousedown', { bubbles: true, cancelable: true })
    el.dispatchEvent(ev)
    return ev
  }

  it('hands focus and selection back on a menu-item press while focus is parked on a field', () => {
    const div = editor('Hello world')
    const input = document.createElement('input')
    const item = document.createElement('button')
    zoneWith(input, item)
    const keeper = createSelectionKeeper('test-kept')
    const commit = vi.fn()
    const off = installKeepEditGuard({ editor: () => div, keeper, commit })
    div.focus()
    select(div, 0, 5)
    keeper.remember(div) // what the editor's blur does
    input.focus()
    window.getSelection()!.removeAllRanges()
    const ev = press(item)
    expect(ev.defaultPrevented).toBe(true)
    expect(commit).not.toHaveBeenCalled()
    expect(document.activeElement).toBe(div)
    expect(window.getSelection()!.toString()).toBe('Hello')
    off()
    expect(keeper.peek()).toBeNull()
  })

  it('a press on a plain zone control keeps focus; a field press is left alone', () => {
    const div = editor('Hello')
    const btn = document.createElement('button')
    const input = document.createElement('input')
    zoneWith(btn, input)
    const keeper = createSelectionKeeper('test-kept')
    const off = installKeepEditGuard({ editor: () => div, keeper, commit: vi.fn() })
    div.focus()
    expect(press(btn).defaultPrevented).toBe(true)
    expect(press(input).defaultPrevented).toBe(false)
    off()
  })

  it('commits only on a press outside every zone while focus is parked', () => {
    const div = editor('Hello')
    const input = document.createElement('input')
    zoneWith(input)
    const outside = document.createElement('div')
    document.body.appendChild(outside)
    const commit = vi.fn()
    const off = installKeepEditGuard({
      editor: () => div,
      keeper: createSelectionKeeper('test-kept'),
      commit,
    })
    div.focus()
    press(outside) // focus still in the editor: its own blur handles this
    expect(commit).not.toHaveBeenCalled()
    input.focus()
    press(outside)
    expect(commit).toHaveBeenCalledTimes(1)
    off()
  })
})
