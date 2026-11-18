import React, { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { beforeEach, afterEach, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem } from '../src/shared/ipc'

vi.mock('react-konva', async () => {
  const React = await import('react')
  const box = ({ children }: { children?: React.ReactNode }) =>
    React.createElement('div', null, children)
  return {
    Stage: React.forwardRef(({ children }: { children?: React.ReactNode }, ref) => {
      React.useImperativeHandle(ref, () => ({ draw: () => {} }))
      return React.createElement('div', null, children)
    }),
    Layer: box,
    Group: box,
    Rect: () => null,
  }
})
vi.mock('../src/renderer/NodeBody', () => ({
  StaticNode: ({ node }: { node: { sourceId: string } }) =>
    React.createElement('div', { 'data-visible-node': node.sourceId }),
}))
vi.mock('../src/renderer/konva-adapter', () => ({ fillToKonva: () => ({}) }))
vi.mock('../src/renderer/i18n/locale', () => ({ useI18n: () => ({ t: (key: string) => key }) }))
import { AnimPreviewOverlay } from '../src/renderer/components/AnimatedSlide'

let root: Root
let host: HTMLDivElement
const items: AnimationItem[] = Array.from({ length: 9 }, (_, i) => ({
  sourceId: `red-${i + 1}`,
  targetName: `Step ${i + 1}`,
  effect: 'appear',
  trigger: 'onClick',
  durationMs: 0,
  delayMs: 0,
  after: 'hideOnNextClick',
}))
const slide = {
  widthPx: 960,
  heightPx: 540,
  scale: 1,
  background: { kind: 'solid', color: '#ffffff' },
  nodes: items.map((a) => ({
    id: a.sourceId,
    sourceId: a.sourceId,
    type: 'shape',
    box: { x: 0, y: 0, w: 100, h: 100 },
  })),
} as RenderSlide
const visible = () =>
  [...host.querySelectorAll('[data-visible-node]')].map((n) => n.getAttribute('data-visible-node'))
beforeEach(() => {
  Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
  vi.useFakeTimers()
  host = document.createElement('div')
  document.body.appendChild(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.useRealTimers()
})
function render(list: AnimationItem[], done: () => void) {
  act(() =>
    root.render(
      React.createElement(AnimPreviewOverlay, {
        slide,
        items: list,
        images: new Map(),
        onDone: done,
      }),
    ),
  )
}
it('automatically previews all nine instant red frames one at a time without parent rerenders', () => {
  const done = vi.fn()
  render(items, done)
  expect(visible()).toEqual([])
  for (let i = 1; i <= 9; i++) {
    act(() => vi.advanceTimersByTime(350))
    expect(visible()).toEqual([`red-${i}`])
    expect(done).not.toHaveBeenCalled()
  }
  act(() => vi.advanceTimersByTime(599))
  expect(done).not.toHaveBeenCalled()
  act(() => vi.advanceTimersByTime(1))
  expect(done).toHaveBeenCalledTimes(1)
})
it('restarts the pause when animation metadata changes during a pending preview', () => {
  const done = vi.fn()
  render(items, done)
  act(() => vi.advanceTimersByTime(200))
  render(
    items.map((a) => ({ ...a })),
    done,
  )
  act(() => vi.advanceTimersByTime(150))
  expect(visible()).toEqual([])
  act(() => vi.advanceTimersByTime(200))
  expect(visible()).toEqual(['red-1'])
})
