import React, { act, useLayoutEffect } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  useAnimPlayer,
  type AnimPlayer,
  type AnimPageSource,
} from '../src/renderer/use-anim-player'
import type { AnimationItem } from '../src/shared/ipc'

const item = (
  sourceId: string,
  effect: AnimationItem['effect'],
  rest: Partial<AnimationItem> = {},
): AnimationItem => ({
  sourceId,
  effect,
  targetName: sourceId,
  trigger: 'onClick',
  durationMs: 0,
  delayMs: 0,
  ...rest,
})
let root: Root
let host: HTMLDivElement
let player: AnimPlayer
let frame: FrameRequestCallback
function Harness() {
  player = useAnimPlayer(540, 960)
  return null
}
const commits: Array<{ key: string | number; played: number; hidden: boolean | undefined }> = []
function PaintedFrame({ source, player: p }: { source: AnimPageSource; player: AnimPlayer }) {
  useLayoutEffect(() => {
    commits.push({ key: source.key, played: p.played, hidden: p.states.get('shared-id')?.hidden })
  })
  return null
}
function PageHarness({ source }: { source: AnimPageSource }) {
  player = useAnimPlayer(540, 960, source)
  return player.ready ? React.createElement(PaintedFrame, { source, player }) : null
}
beforeEach(() => {
  Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
  commits.length = 0
  vi.stubGlobal(
    'requestAnimationFrame',
    vi.fn((callback: FrameRequestCallback) => {
      frame = callback
      return 1
    }),
  )
  vi.stubGlobal('cancelAnimationFrame', vi.fn())
  host = document.createElement('div')
  document.body.appendChild(host)
  root = createRoot(host)
  act(() => root.render(React.createElement(Harness)))
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

describe('presenter/audience animation cursor', () => {
  it('resets every revisit before the new page commits, including a jump to the same page', () => {
    const items = [item('shared-id', 'appear'), item('second', 'appear')]
    const render = (key: string) =>
      act(() => root.render(React.createElement(PageHarness, { source: { key, items } })))
    render('9:1')
    act(() => {
      player.advance()
      player.advance()
    })
    expect(player.pending).toBe(false)
    render('10:2')
    render('9:3')
    expect(commits.filter((c) => c.key === '9:3')).toEqual([
      { key: '9:3', played: 0, hidden: true },
    ])
    act(() => {
      player.advance()
    })
    render('9:4')
    expect(commits.filter((c) => c.key === '9:4')).toEqual([
      { key: '9:4', played: 0, hidden: true },
    ])
  })
  it('keeps the frame unmounted while animation metadata is pending', () => {
    act(() => root.render(React.createElement(PageHarness, { source: { key: 9, items: null } })))
    expect(player.ready).toBe(false)
    expect(commits).toEqual([])
    act(() =>
      root.render(
        React.createElement(PageHarness, {
          source: { key: 9, items: [item('shared-id', 'appear')] },
        }),
      ),
    )
    expect(commits).toEqual([{ key: 9, played: 0, hidden: true }])
  })
  it('applies audience cursors before canvas commit and ignores a retired page animation tick', () => {
    const items = [item('shared-id', 'fade', { durationMs: 500 })]
    const render = (key: number, played: number, playing: boolean) =>
      act(() =>
        root.render(
          React.createElement(PageHarness, { source: { key, items, cursor: { played, playing } } }),
        ),
      )
    render(1, 0, true)
    const retiredFrame = frame
    render(2, 0, false)
    expect(commits.filter((c) => c.key === 2)).toEqual([{ key: 2, played: 0, hidden: true }])
    act(() => retiredFrame(performance.now() + 1000))
    expect(player.played).toBe(0)
    render(2, 1, false)
    expect(commits.at(-1)).toEqual({ key: 2, played: 1, hidden: false })
    const epoch = player.epoch
    render(2, 1, false)
    expect(player.epoch).toBe(epoch)
  })
  it('completes instant reveal/hide steps without waiting for a background animation frame', () => {
    act(() =>
      player.load(
        [
          item('step1', 'appear'),
          item('step2', 'appear'),
          item('step1', 'disappear', { trigger: 'withPrev' }),
        ],
        'fresh',
      ),
    )
    expect(player.states.get('step1')!.hidden).toBe(true)
    act(() => {
      expect(player.advance()).toBe(true)
    })
    expect(player.played).toBe(1)
    expect(player.playing).toBe(false)
    expect(player.states.get('step1')!.hidden).toBe(false)
    act(() => {
      expect(player.advance()).toBe(true)
    })
    expect(player.played).toBe(2)
    expect(player.states.get('step1')!.hidden).toBe(true)
    expect(player.states.get('step2')!.hidden).toBe(false)
    expect(player.advance()).toBe(false)
  })
  it('handles consecutive instant input before React paints and mirrors the same absolute cursor', () => {
    const items = [item('a', 'appear'), item('b', 'appear')]
    act(() => player.load(items, 'fresh'))
    act(() => {
      player.advance()
      player.advance()
    })
    expect(player.played).toBe(2)
    expect(player.pending).toBe(false)
    act(() => player.seek(items, 1, true))
    expect(player.played).toBe(2)
    expect(player.playing).toBe(false)
    act(() => player.seek(items, 0, false))
    expect(player.played).toBe(0)
    expect(player.states.get('a')!.hidden).toBe(true)
  })
  it('settles automatic instant entrance on load while keeping click steps pending', () => {
    act(() =>
      player.load(
        [item('auto', 'appear', { trigger: 'withPrev' }), item('click', 'appear')],
        'fresh',
      ),
    )
    expect(player.played).toBe(1)
    expect(player.playing).toBe(false)
    expect(player.states.get('auto')!.hidden).toBe(false)
    expect(player.states.get('click')!.hidden).toBe(true)
  })
  it('retains timed playback and honors delays on instant effects', () => {
    const now = vi.spyOn(performance, 'now').mockReturnValue(100)
    act(() => player.load([item('timed', 'fade', { durationMs: 500 })], 'fresh'))
    act(() => {
      player.advance()
    })
    expect(player.playing).toBe(true)
    now.mockReturnValue(350)
    act(() => frame(350))
    expect(player.states.get('timed')!.opacity).toBeGreaterThan(0)
    expect(player.states.get('timed')!.opacity).toBeLessThan(1)
    now.mockReturnValue(600)
    act(() => frame(600))
    expect(player.played).toBe(1)
    act(() => player.load([item('delayed', 'appear', { delayMs: 300 })], 'fresh'))
    act(() => {
      player.advance()
    })
    expect(player.playing).toBe(true)
    expect(player.played).toBe(0)
  })
})
