import React, { act, useLayoutEffect } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem, ShowSyncState } from '../src/shared/ipc'
import type { NodeAnimState } from '../src/renderer/animation-play'

const frames: Array<{ slide: RenderSlide; hidden: boolean | undefined }> = []
vi.mock('../src/renderer/components/AnimatedSlide', async () => {
  const player = await import('../src/renderer/use-anim-player')
  return {
    useAnimPlayer: player.useAnimPlayer,
    AnimatedSlideStage: ({
      slide,
      states,
    }: {
      slide: RenderSlide
      states: Map<string, NodeAnimState>
    }) => {
      useLayoutEffect(() => {
        frames.push({ slide, hidden: states.get('shared-id')?.hidden })
      })
      return null
    },
  }
})
vi.mock('../src/renderer/doc-fonts', () => ({ syncPrivateFonts: async () => {} }))
vi.mock('../src/renderer/components/MorphStage', () => ({ MorphStage: () => null }))
vi.mock('../src/renderer/components/ShowCamera', () => ({ CameraBubble: () => null }))
vi.mock('../src/renderer/components/ShowInk', () => ({
  InkBoard: class {
    clear() {}
  },
  InkLayer: () => null,
  applyInkEvent: () => {},
  useInkPointer: () => ({}),
  DEFAULT_PEN_COLOR: '#000000',
}))

import { AudienceView } from '../src/renderer/components/AudienceView'

it('never commits a new audience page with missing metadata or the preceding page cursor', async () => {
  Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
  const slide = (): RenderSlide => ({
    widthPx: 960,
    heightPx: 540,
    scale: 1,
    nodes: [],
    background: { kind: 'solid', color: '#ffffff' },
  })
  const slides = [slide(), slide()]
  const items: AnimationItem[] = [
    {
      sourceId: 'shared-id',
      targetName: 'entrance',
      effect: 'appear',
      trigger: 'onClick',
      durationMs: 0,
      delayMs: 0,
    },
  ]
  const pending: Array<(items: AnimationItem[]) => void> = []
  let receiveSync: (s: ShowSyncState) => void = () => {}
  const sync = (idx: number, played: number, visit: number): ShowSyncState => ({
    idx,
    played,
    visit,
    playing: false,
    fresh: false,
    ended: false,
    black: false,
  })
  const originalApi = window.slidesApi
  window.slidesApi = {
    getRenderSlides: async () => slides,
    getTransition: async () => 'none',
    getShapeKeys: async () => [],
    getAnimations: () => new Promise<AnimationItem[]>((resolve) => pending.push(resolve)),
    onShowSync: (receive: typeof receiveSync) => {
      receiveSync = receive
      return () => {}
    },
    onShowInk: () => () => {},
    onDeckChanged: () => () => {},
    audienceReady: async () => sync(0, 1, 0),
  } as unknown as typeof window.slidesApi
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  try {
    await act(async () => {
      root.render(React.createElement(AudienceView))
    })
    expect(pending).toHaveLength(2)
    expect(frames).toEqual([])
    await act(async () => {
      pending.forEach((resolve) => resolve(items))
    })
    expect(frames.at(-1)).toEqual({ slide: slides[0], hidden: false })
    const before = frames.length
    act(() => receiveSync(sync(1, 0, 1)))
    expect(frames.slice(before).length).toBeGreaterThan(0)
    expect(frames.slice(before).every((f) => f.slide === slides[1] && f.hidden === true)).toBe(true)
    act(() => receiveSync(sync(1, 1, 1)))
    expect(frames.at(-1)).toEqual({ slide: slides[1], hidden: false })
    const returning = frames.length
    act(() => receiveSync(sync(0, 0, 2)))
    expect(frames.slice(returning).every((f) => f.slide === slides[0] && f.hidden === true)).toBe(
      true,
    )
  } finally {
    act(() => root.unmount())
    host.remove()
    window.slidesApi = originalApi
    vi.restoreAllMocks()
  }
})
