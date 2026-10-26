// Regression: using the ribbon / format pane while text is selected must not end the edit or lose
// the selection (PowerPoint keeps both). The old global press listener committed the edit whenever
// focus was parked on a ribbon field (font box, size box) and the next press landed on a menu item
// or button that was not tagged keep-edit.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import {
  TextEditOverlay,
  classifyEditPress,
  editorIsDirty,
  forgetEditSelection,
  rangeFromOffsets,
  rememberEditSelection,
  restoreEditSelection,
  selectionOffsets,
} from '../src/renderer/TextEditOverlay'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function fakeNode(over: Partial<ShapeRenderNode> = {}): ShapeRenderNode {
  return {
    id: 'n1',
    type: 'shape',
    sourceId: 'src1',
    box: { x: 0, y: 0, w: 300, h: 60, rotationDeg: 0 },
    fill: { kind: 'none' },
    ...over,
  } as unknown as ShapeRenderNode
}

describe('classifyEditPress', () => {
  it('splits editor frame, neutral zone, neutral field and outside', () => {
    const frame = document.createElement('div')
    const inner = document.createElement('span')
    frame.appendChild(inner)
    const zone = document.createElement('div')
    zone.setAttribute('data-keep-edit', '')
    const btn = document.createElement('button')
    const label = document.createElement('label')
    const input = document.createElement('input')
    const nested = document.createElement('span')
    label.appendChild(nested)
    zone.append(btn, label, input)
    const other = document.createElement('div')
    document.body.append(frame, zone, other)
    expect(classifyEditPress(frame, inner)).toBe('in-editor')
    expect(classifyEditPress(frame, btn)).toBe('neutral')
    // a label (blank panel text) is not a focus field: the editor must keep focus
    expect(classifyEditPress(frame, nested)).toBe('neutral')
    expect(classifyEditPress(frame, input)).toBe('neutral-field')
    expect(classifyEditPress(frame, other)).toBe('outside')
    expect(classifyEditPress(frame, null)).toBe('outside')
    frame.remove()
    zone.remove()
    other.remove()
  })
})

describe('selection offsets', () => {
  it('round-trips a selection across a rebuilt DOM by character offsets', () => {
    const a = document.createElement('div')
    a.innerHTML = '<div>Hello <b>big</b> world</div><div>second line</div>'
    document.body.appendChild(a)
    const range = document.createRange()
    range.setStart(a.querySelector('b')!.firstChild!, 1)
    range.setEnd(a.lastElementChild!.firstChild!, 6)
    const sel = window.getSelection()!
    sel.removeAllRanges()
    sel.addRange(range)
    const off = selectionOffsets(a)!
    // "Hello big world" is 15 characters, then 6 into the second paragraph
    expect(off).toEqual({ start: 7, end: 21 })
    // rebuild with a different inline structure; same text
    a.innerHTML = '<div>Hello big world</div><div>second line</div>'
    const back = rangeFromOffsets(a, off.start, off.end)
    expect(back.toString()).toBe('ig worldsecond')
    a.remove()
  })

  it('clamps offsets past the text and handles an empty root', () => {
    const a = document.createElement('div')
    document.body.appendChild(a)
    const r = rangeFromOffsets(a, 3, 9)
    expect(r.collapsed).toBe(true)
    a.textContent = 'abc'
    expect(rangeFromOffsets(a, 1, 99).toString()).toBe('bc')
    a.remove()
  })

  it('selectionOffsets is null for a selection outside the root', () => {
    const a = document.createElement('div')
    a.textContent = 'inside'
    const b = document.createElement('div')
    b.textContent = 'outside'
    document.body.append(a, b)
    const range = document.createRange()
    range.selectNodeContents(b)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(range)
    expect(selectionOffsets(a)).toBeNull()
    a.remove()
    b.remove()
  })
})

describe('editorIsDirty', () => {
  it('compares the live extraction with the session snapshot', () => {
    const paras = [{ runs: [{ text: 'a' }] }] as never
    expect(editorIsDirty(JSON.stringify(paras), paras)).toBe(false)
    expect(editorIsDirty(JSON.stringify(paras), [{ runs: [{ text: 'ab' }] }] as never)).toBe(true)
  })
})

describe('kept selection handoff', () => {
  it('rememberEditSelection snapshots the range when focus leaves for a keep-edit field', () => {
    const div = document.createElement('div')
    div.contentEditable = 'true'
    div.tabIndex = 0
    Object.defineProperty(div, 'isContentEditable', { value: true })
    div.textContent = 'AAABBB'
    document.body.appendChild(div)
    div.focus()
    const range = document.createRange()
    range.setStart(div.firstChild!, 1)
    range.setEnd(div.firstChild!, 4)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(range)
    rememberEditSelection(div)
    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    window.getSelection()!.removeAllRanges()
    expect(restoreEditSelection()).toBe(true)
    expect(document.activeElement).toBe(div)
    expect(window.getSelection()!.toString()).toBe('AAB')
    forgetEditSelection()
    expect(restoreEditSelection()).toBe(false)
    div.remove()
    input.remove()
  })
})

describe('TextEditOverlay in a live ribbon session', () => {
  let host: HTMLDivElement
  let root: Root
  beforeEach(() => {
    host = document.createElement('div')
    document.body.appendChild(host)
    root = createRoot(host)
  })
  afterEach(() => {
    act(() => root.unmount())
    host.remove()
    forgetEditSelection()
  })

  const mount = (node: ShapeRenderNode, onCommit = vi.fn(), onCancel = vi.fn()) => {
    const zone = document.createElement('div')
    zone.setAttribute('data-keep-edit', '')
    const input = document.createElement('input')
    const menuItem = document.createElement('button')
    zone.append(input, menuItem)
    const outside = document.createElement('div')
    document.body.append(zone, outside)
    const render = (n: ShapeRenderNode) =>
      act(() =>
        root.render(createElement(TextEditOverlay, { node: n, scale: 1, onCommit, onCancel })),
      )
    render(node)
    const editor = host.querySelector('.slide-text-editor') as HTMLElement
    Object.defineProperty(editor, 'isContentEditable', { value: true })
    return { zone, input, menuItem, outside, editor, onCommit, onCancel, render }
  }

  const press = (el: Element) =>
    act(() => {
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    })

  it('a press on a menu item while focus is parked on a ribbon field hands focus and the selection back', () => {
    const { editor, input, menuItem, zone, outside, onCommit, onCancel } = mount(fakeNode())
    editor.textContent = 'Hello world'
    editor.focus()
    const range = document.createRange()
    range.setStart(editor.firstChild!, 0)
    range.setEnd(editor.firstChild!, 5)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(range)
    // press on the size box: a field, the editor blurs into it
    press(input)
    act(() => input.focus())
    window.getSelection()!.removeAllRanges()
    expect(document.activeElement).toBe(input)
    // press on a menu item (not a field): the old listener committed here
    press(menuItem)
    expect(onCommit).not.toHaveBeenCalled()
    expect(onCancel).not.toHaveBeenCalled()
    expect(document.activeElement).toBe(editor)
    expect(window.getSelection()!.toString()).toBe('Hello')
    zone.remove()
    outside.remove()
  })

  it('a press outside every keep-edit zone still commits while focus is parked', () => {
    const { editor, input, outside, zone, onCommit, onCancel } = mount(fakeNode())
    editor.textContent = 'Hello world'
    editor.focus()
    act(() => input.focus())
    press(outside)
    expect(onCommit.mock.calls.length + onCancel.mock.calls.length).toBe(1)
    zone.remove()
    outside.remove()
  })

  it('a model refresh keeps the typed text and selection of a touched editor', () => {
    const { editor, render, zone, outside } = mount(fakeNode())
    editor.focus()
    editor.textContent = 'typed by the user'
    const range = document.createRange()
    range.setStart(editor.firstChild!, 6)
    range.setEnd(editor.firstChild!, 8)
    window.getSelection()!.removeAllRanges()
    window.getSelection()!.addRange(range)
    // same text box, new node object (a format pane edit re-laid it out)
    render(fakeNode({ box: { x: 0, y: 0, w: 300, h: 90, rotationDeg: 0 } as never }))
    expect(editor.textContent).toBe('typed by the user')
    expect(window.getSelection()!.toString()).toBe('by')
    zone.remove()
    outside.remove()
  })
})
