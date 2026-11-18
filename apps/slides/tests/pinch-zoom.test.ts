import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

// react-konva's node entry requires the native 'canvas' package; nothing here draws
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import { App } from '../src/renderer/App'

class FakeResizeObserver {
  static instances: FakeResizeObserver[] = []
  targets = new Set<Element>()
  constructor(private cb: ResizeObserverCallback) {
    FakeResizeObserver.instances.push(this)
  }
  observe(el: Element) {
    this.targets.add(el)
    this.cb([{ target: el } as ResizeObserverEntry], this as unknown as ResizeObserver)
  }
  unobserve(el: Element) {
    this.targets.delete(el)
  }
  disconnect() {
    this.targets.clear()
  }
}

let stageSize = { w: 1336, h: 800 }
let outerSize = { w: 1336, h: 800 }

const blankSlide = () => ({
  widthPx: 1280,
  heightPx: 720,
  scale: 1,
  nodes: [],
  background: { kind: 'color', color: '#ffffff' },
})

function makeSlidesApi() {
  const explicit: Record<string, unknown> = {
    consumePendingOpen: () => Promise.resolve(null),
    newBlank: () => Promise.resolve({ path: '', slides: [blankSlide()], defaultFont: 'Arial' }),
    isDirty: () => Promise.resolve(false),
    getRecentFiles: () => Promise.resolve([]),
    getAiSettings: () => Promise.resolve(null),
    getSections: () => Promise.resolve([]),
    getComments: () => Promise.resolve([]),
    getNotes: () => Promise.resolve(''),
    getAnimations: () => Promise.resolve([]),
    getTransition: () => Promise.resolve(null),
    getShapeKeys: () => Promise.resolve([]),
    getSlideLinks: () => Promise.resolve([]),
    getRunLinks: () => Promise.resolve([]),
    listFonts: () => Promise.resolve([]),
  }
  const fallbacks = new Map<string, unknown>()
  return new Proxy(explicit, {
    get(target, prop: string) {
      if (prop in target) return target[prop]
      if (!fallbacks.has(prop)) {
        fallbacks.set(
          prop,
          prop.startsWith('on')
            ? () => () => {}
            : () => Promise.resolve(null),
        )
      }
      return fallbacks.get(prop)
    },
  })
}

async function settle() {
  for (let i = 0; i < 8; i++) {
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
  }
}

async function flushZoomPreview() {
  await act(async () => {
    await new Promise((r) => requestAnimationFrame(() => r(null)))
    await new Promise((r) => setTimeout(r, 280))
  })
}

function stageZoom(container: HTMLElement): number {
  const el = container.querySelector<HTMLElement>('.stage-scale')
  expect(el, 'stage-scale should be mounted').not.toBeNull()
  const m = /scale\(([\d.]+)\)/.exec(el!.style.transform)
  expect(m, `transform should carry a scale(): ${el!.style.transform}`).not.toBeNull()
  return Number(m![1])
}

let root: Root | null = null
let container: HTMLElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  vi.stubGlobal('ResizeObserver', FakeResizeObserver)
  window.matchMedia ??= (query: string) =>
    ({
      matches: false,
      media: query,
      addEventListener: () => {},
      removeEventListener: () => {},
      addListener: () => {},
      removeListener: () => {},
      onchange: null,
      dispatchEvent: () => false,
    }) as MediaQueryList
  Element.prototype.scrollTo ??= () => {}
  Element.prototype.scrollIntoView ??= () => {}
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get: () => stageSize.w,
  })
  Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
    configurable: true,
    get: () => stageSize.h,
  })
  Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('stage-wrap') ? outerSize.w : 0
    },
  })
  Object.defineProperty(HTMLElement.prototype, 'offsetHeight', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('stage-wrap') ? outerSize.h : 0
    },
  })
  ;(window as unknown as { slidesApi: unknown }).slidesApi = makeSlidesApi()
})

afterEach(() => {
  act(() => root?.unmount())
  container?.remove()
  root = null
  container = null
  FakeResizeObserver.instances.length = 0
})

async function bootApp(): Promise<HTMLElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => root!.render(createElement(App)))
  await settle()
  const wrap = container.querySelector<HTMLElement>('.stage-wrap')
  expect(wrap, 'editor should boot into a blank deck').not.toBeNull()
  expect(stageZoom(container!)).toBeCloseTo(1, 5)
  return wrap!
}

describe('trackpad pinch-to-zoom continuous gesture scaling', () => {
  it('scales smoothly and continuously with gesture amplitude instead of discrete chunks', async () => {
    const wrap = await bootApp()
    const zoomVal = container!.querySelector<HTMLElement>('.zoom-value')!
    expect(zoomVal.textContent).toBe('100%')

    // Trackpad pinch zoom-in: small gesture amplitude deltaY = -5
    act(() => {
      wrap.dispatchEvent(
        new WheelEvent('wheel', {
          ctrlKey: true,
          deltaY: -5,
          clientX: 500,
          clientY: 400,
          bubbles: true,
          cancelable: true,
        }),
      )
    })

    // Advance rAF so preview transform lands
    await act(async () => {
      await new Promise((r) => requestAnimationFrame(() => r(null)))
    })

    // Scale should have increased by exp(5 * 0.0075) ≈ 1.0382 (3.8% continuous change)
    const scaleSmall = stageZoom(container!)
    expect(scaleSmall).toBeCloseTo(Math.exp(5 * 0.0075), 3)
    expect(zoomVal.textContent).toBe('104%')

    // Additional tick with larger delta: deltaY = -20
    act(() => {
      wrap.dispatchEvent(
        new WheelEvent('wheel', {
          ctrlKey: true,
          deltaY: -20,
          clientX: 500,
          clientY: 400,
          bubbles: true,
          cancelable: true,
        }),
      )
    })
    await act(async () => {
      await new Promise((r) => requestAnimationFrame(() => r(null)))
    })

    const scaleLarger = stageZoom(container!)
    expect(scaleLarger).toBeCloseTo(scaleSmall * Math.exp(20 * 0.0075), 3)
    expect(zoomVal.textContent).toBe('121%')

    // Flush debounce to ensure React state commits the exact continuous value
    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(scaleLarger, 3)
  })

  it('scales continuously on zoom-out and updates status bar in real time', async () => {
    const wrap = await bootApp()
    const zoomVal = container!.querySelector<HTMLElement>('.zoom-value')!

    // Trackpad pinch zoom-out: deltaY = +15
    act(() => {
      wrap.dispatchEvent(
        new WheelEvent('wheel', {
          ctrlKey: true,
          deltaY: 15,
          clientX: 500,
          clientY: 400,
          bubbles: true,
          cancelable: true,
        }),
      )
    })
    await act(async () => {
      await new Promise((r) => requestAnimationFrame(() => r(null)))
    })

    const scaleOut = stageZoom(container!)
    expect(scaleOut).toBeCloseTo(Math.exp(-15 * 0.0075), 3)
    expect(zoomVal.textContent).toBe('89%')

    await flushZoomPreview()
    expect(stageZoom(container!)).toBeCloseTo(scaleOut, 3)
  })
})
