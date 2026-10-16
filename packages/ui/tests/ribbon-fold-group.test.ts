/** @vitest-environment jsdom */
import { act, createElement, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { RibbonFoldGroup } from '../src/RibbonFoldGroup'

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

function NestedPicker() {
  const [open, setOpen] = useState(false)
  return createElement(
    'div',
    null,
    createElement('button', { type: 'button', onClick: () => setOpen((v) => !v) }, 'Symbol'),
    open &&
      createElement(
        'div',
        { 'data-rb-panel': '' },
        createElement('button', { type: 'button', 'data-rb-keep-open': '' }, 'Category'),
        createElement('button', { type: 'button', onClick: () => setOpen(false) }, '≤'),
      ),
  )
}

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.useFakeTimers()
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  act(() =>
    root.render(
      createElement(RibbonFoldGroup, {
        label: 'Symbols',
        icon: '◇',
        priority: 1,
        children: createElement(NestedPicker),
      }),
    ),
  )
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.useRealTimers()
  vi.unstubAllGlobals()
})

const folded = () => host.querySelector('.rb-fold-btn') as HTMLButtonElement
const command = (label: string) =>
  [...host.querySelectorAll('button')].find(
    (button) => button.textContent === label,
  ) as HTMLButtonElement

it('keeps the folded group open for nested launchers and category steps, then closes on a leaf', () => {
  act(() => folded().click())
  act(() => command('Symbol').click())
  act(() => vi.runAllTimers())
  expect(folded().getAttribute('aria-expanded')).toBe('true')
  expect(host.querySelector('[data-rb-panel]')).toBeTruthy()
  act(() => command('Category').click())
  expect(folded().getAttribute('aria-expanded')).toBe('true')
  act(() => command('≤').click())
  expect(folded().getAttribute('aria-expanded')).toBe('false')
})

it('closes for an outside press and Escape', () => {
  act(() => folded().click())
  act(() => document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })))
  expect(folded().getAttribute('aria-expanded')).toBe('false')
  act(() => folded().click())
  act(() => document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })))
  expect(folded().getAttribute('aria-expanded')).toBe('false')
})

it('optionally prevents launcher pointer focus without blocking the folded panel click', () => {
  const editor = document.createElement('textarea')
  document.body.append(editor)
  editor.focus()
  const ordinaryDown = new MouseEvent('pointerdown', { bubbles: true, cancelable: true })
  act(() => folded().dispatchEvent(ordinaryDown))
  expect(ordinaryDown.defaultPrevented).toBe(false)

  act(() =>
    root.render(
      createElement(RibbonFoldGroup, {
        label: 'Symbols',
        icon: '◇',
        priority: 1,
        preserveEditorFocus: true,
        children: createElement(NestedPicker),
      }),
    ),
  )
  const pointerDown = new MouseEvent('pointerdown', { bubbles: true, cancelable: true })
  const mouseDown = new MouseEvent('mousedown', { bubbles: true, cancelable: true })
  act(() => {
    folded().dispatchEvent(pointerDown)
    folded().dispatchEvent(mouseDown)
  })
  expect(pointerDown.defaultPrevented).toBe(true)
  expect(mouseDown.defaultPrevented).toBe(true)
  expect(document.activeElement).toBe(editor)
  act(() => folded().click())
  expect(folded().getAttribute('aria-expanded')).toBe('true')
  editor.remove()
})
