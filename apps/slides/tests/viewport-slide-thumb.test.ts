import { createElement, act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import { ViewportSlideThumb } from '../src/renderer/ViewportSlideThumb'

vi.mock('../src/renderer/SlideThumb', () => ({ SlideThumb: () => createElement('canvas') }))

let notify: IntersectionObserverCallback
const disconnect = vi.fn()
vi.stubGlobal(
  'IntersectionObserver',
  class {
    constructor(callback: IntersectionObserverCallback) {
      notify = callback
    }
    observe() {}
    disconnect = disconnect
  },
)
;(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true
const container = document.createElement('div')
const root = createRoot(container)
const slide = { widthPx: 1280, heightPx: 720 } as RenderSlide
const images = new Map()

afterEach(() => {
  act(() => root.render(null))
  disconnect.mockClear()
})

it('preserves scroll geometry and releases offscreen canvases, including after an edit', () => {
  const render = (current = false) =>
    act(() =>
      root.render(createElement(ViewportSlideThumb, { slide, images, width: 160, current })),
    )
  render()
  expect(container.firstElementChild?.getAttribute('style')).toContain('height: 90px')
  expect(container.querySelector('canvas')).toBeNull()
  act(() =>
    notify([{ isIntersecting: true } as IntersectionObserverEntry], {} as IntersectionObserver),
  )
  expect(container.querySelector('canvas')).not.toBeNull()
  act(() =>
    notify([{ isIntersecting: false } as IntersectionObserverEntry], {} as IntersectionObserver),
  )
  expect(container.querySelector('canvas')).toBeNull()
  render(true)
  expect(container.querySelector('canvas')).not.toBeNull()
  render(false)
  expect(container.querySelector('canvas')).toBeNull()
  act(() => root.render(null))
  expect(disconnect).toHaveBeenCalledTimes(1)
})
