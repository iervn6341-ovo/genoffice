import React, { act, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem } from '../src/shared/ipc'
import { AnimationPane } from '../src/renderer/components/AnimationPane'
import { AnimationOrderOverlay } from '../src/renderer/components/AnimationOrderOverlay'
import { editedAnimation, type AnimationEditPatch } from '../src/renderer/animation-actions'
import { LocaleProvider } from '../src/renderer/i18n/locale'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
Object.assign(window, { slidesApi: { onLanguageChanged: () => () => {} } })
const slide = {
  nodes: [{ id: 'a', sourceId: 'a', type: 'shape', box: { x: 200, y: 100, w: 80, h: 60 } }],
  widthPx: 960,
  heightPx: 540,
  scale: 1,
  background: { kind: 'solid', color: '#ffffff' },
} as RenderSlide
const initial: AnimationItem[] = [
  {
    sourceId: 'a',
    targetName: 'Card',
    effect: 'appear',
    trigger: 'onClick',
    durationMs: 0,
    delayMs: 0,
  },
  {
    sourceId: 'a',
    targetName: 'Card',
    effect: 'fadeOut',
    trigger: 'onClick',
    durationMs: 500,
    delayMs: 0,
  },
]
const patches = vi.fn()
const moved = vi.fn()
const deleted = vi.fn()
let current: AnimationItem[]
const host = document.createElement('div')
document.body.appendChild(host)
const root = createRoot(host)
const scrolled = vi.fn()
Object.defineProperty(HTMLElement.prototype, 'scrollIntoView', {
  value: scrolled,
  configurable: true,
})
function Harness() {
  const [selected, setSelected] = useState(-1)
  const [items, setItems] = useState(initial)
  current = items
  const onPatch = (index: number, patch: AnimationEditPatch) => {
    patches(index, patch)
    setItems((list) => list.map((item, i) => (i === index ? editedAnimation(item, patch) : item)))
  }
  return React.createElement(LocaleProvider, {
    initial: 'zh-TW',
    children: React.createElement(
      React.Fragment,
      null,
      React.createElement(AnimationOrderOverlay, {
        slide,
        items,
        zoom: 1,
        selected,
        onSelect: setSelected,
        onDelete: deleted,
      }),
      React.createElement(AnimationPane, {
        slideIndex: 8,
        items,
        selected,
        onSelect: setSelected,
        onPatch,
        onMove: moved,
        onDelete: deleted,
        onPreview: () => {},
        onCollapse: () => {},
      }),
    ),
  })
}
afterEach(() => {
  act(() => root.render(null))
  patches.mockClear()
  moved.mockClear()
  deleted.mockClear()
  scrolled.mockClear()
})
function pick(label: string, value: string) {
  const select = host.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`)!
  act(() => {
    select.value = value
    select.dispatchEvent(new Event('change', { bubbles: true }))
  })
}

it('clicking a number selects and scrolls its exact effect, and row selection highlights the matching number', () => {
  act(() => root.render(React.createElement(Harness)))
  act(() => host.querySelector<HTMLButtonElement>('[data-testid="animation-order-1"]')!.click())
  expect(host.querySelector('[data-testid="animation-row-1"]')!.getAttribute('aria-selected')).toBe(
    'true',
  )
  expect(
    host.querySelector('[data-testid="animation-order-1"]')!.getAttribute('aria-pressed'),
  ).toBe('true')
  expect(scrolled).toHaveBeenCalled()
  expect(host.querySelector<HTMLSelectElement>('select[aria-label="動畫"]')!.value).toBe('fadeOut')
  act(() => host.querySelector<HTMLElement>('[data-testid="animation-row-0"]')!.click())
  expect(
    host.querySelector('[data-testid="animation-order-0"]')!.getAttribute('aria-pressed'),
  ).toBe('true')
  expect(host.querySelector<HTMLInputElement>('input[aria-label="期間："]')!.disabled).toBe(true)
  expect(patches).not.toHaveBeenCalled()
})

it('editing the selected effect updates its properties without changing another effect on the same shape', () => {
  act(() => root.render(React.createElement(Harness)))
  act(() => host.querySelector<HTMLElement>('[data-testid="animation-row-0"]')!.click())
  pick('動畫', 'flyIn')
  pick('效果選項', 'left')
  pick('開始：', 'afterPrev')
  pick('動畫結束後：', 'hideOnNextClick')
  const duration = host.querySelector<HTMLInputElement>('input[aria-label="期間："]')!
  act(() => {
    duration.value = '1.25'
    duration.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
  })
  expect(current[0]).toMatchObject({
    effect: 'flyIn',
    direction: 'left',
    trigger: 'afterPrev',
    after: 'hideOnNextClick',
    durationMs: 1250,
  })
  expect(current[1]).toBe(initial[1])
  const row = host.querySelector<HTMLElement>('[data-testid="animation-row-0"]')!
  act(() => row.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true })))
  expect(deleted).toHaveBeenCalledWith(0)
})
