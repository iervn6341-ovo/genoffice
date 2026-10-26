import { afterEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { restoreFocusAfterModal, useModalKeys } from '../src/renderer/components/modal-keys'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function Modal() {
  const keys = useModalKeys(() => {})
  return createElement(
    'div',
    { ref: keys.ref, onKeyDown: keys.onKeyDown, tabIndex: -1 },
    createElement('input', { id: 'first' }),
  )
}

let root: Root | null = null
afterEach(() => {
  act(() => root?.unmount())
  root = null
  document.body.innerHTML = ''
})

function mountEditor(): HTMLElement {
  const ed = document.createElement('div')
  ed.className = 'ProseMirror doc-page'
  ed.tabIndex = 0
  document.body.appendChild(ed)
  return ed
}

describe('modal returns focus to the document', () => {
  it('focus lands on the editor after the dialog unmounts, not on BODY', () => {
    const ed = mountEditor()
    ed.focus()
    const host = document.createElement('div')
    document.body.appendChild(host)
    root = createRoot(host)
    act(() => root!.render(createElement(Modal)))
    expect(document.activeElement?.id).toBe('first')
    act(() => root!.render(null))
    expect(document.activeElement).toBe(ed)
  })

  it('falls back to the editor when the dialog was opened with focus already on BODY', () => {
    const ed = mountEditor()
    ;(document.activeElement as HTMLElement | null)?.blur()
    const host = document.createElement('div')
    document.body.appendChild(host)
    root = createRoot(host)
    act(() => root!.render(createElement(Modal)))
    act(() => root!.render(null))
    expect(document.activeElement).toBe(ed)
  })

  it('does not steal focus a close deliberately moved elsewhere', () => {
    mountEditor()
    const other = document.createElement('input')
    document.body.appendChild(other)
    other.focus()
    expect(document.activeElement).toBe(other)
    restoreFocusAfterModal(null)
    expect(document.activeElement).toBe(other)
  })
})
