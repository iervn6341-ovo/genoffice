import { describe, it, expect } from 'vitest'
import {
  applyVolumeInput,
  bufferedEnd,
  clampMediaTime,
  controlsVisible,
  formatMediaTime,
  mediaFraction,
  toggleMute,
  videoBarLayout,
  volumeLevel,
} from '../src/renderer/show-video-state'
import { claimShowSpace, registerShowVideoSpaceClaimer } from '../src/renderer/show-video-space'

const key = (k: string, mods: Partial<KeyboardEvent> = {}) =>
  ({ key: k, metaKey: false, ctrlKey: false, altKey: false, ...mods }) as KeyboardEvent

describe('formatMediaTime', () => {
  it('formats m:ss and h:mm:ss', () => {
    expect(formatMediaTime(0)).toBe('0:00')
    expect(formatMediaTime(5.9)).toBe('0:05')
    expect(formatMediaTime(65)).toBe('1:05')
    expect(formatMediaTime(3599)).toBe('59:59')
    expect(formatMediaTime(3600)).toBe('1:00:00')
    expect(formatMediaTime(3725)).toBe('1:02:05')
  })
  it('treats unknown values as 0:00', () => {
    expect(formatMediaTime(NaN)).toBe('0:00')
    expect(formatMediaTime(Infinity)).toBe('0:00')
    expect(formatMediaTime(-4)).toBe('0:00')
  })
})

describe('time maths', () => {
  it('clamps and normalises', () => {
    expect(clampMediaTime(12, 10)).toBe(10)
    expect(clampMediaTime(-1, 10)).toBe(0)
    expect(clampMediaTime(3, NaN)).toBe(0)
    expect(mediaFraction(5, 10)).toBe(0.5)
    expect(mediaFraction(15, 10)).toBe(1)
    expect(mediaFraction(5, 0)).toBe(0)
    expect(mediaFraction(5, NaN)).toBe(0)
  })
  it('finds the buffered end containing the playhead', () => {
    const r = (pairs: [number, number][]) => ({
      length: pairs.length,
      start: (i: number) => pairs[i]![0],
      end: (i: number) => pairs[i]![1],
    })
    expect(bufferedEnd(r([[0, 4]]), 2)).toBe(4)
    expect(
      bufferedEnd(
        r([
          [0, 4],
          [6, 9],
        ]),
        7,
      ),
    ).toBe(9)
    expect(bufferedEnd(r([[0, 4]]), 5)).toBe(5)
    expect(bufferedEnd(null, 3)).toBe(3)
  })
})

describe('controls state', () => {
  it('shows only while the pointer is active, or a drag runs', () => {
    expect(controlsVisible({ hovering: false, idle: false, dragging: false })).toBe(false)
    expect(controlsVisible({ hovering: true, idle: false, dragging: false })).toBe(true)
    expect(controlsVisible({ hovering: true, idle: true, dragging: false })).toBe(false)
    expect(controlsVisible({ hovering: true, idle: true, dragging: true })).toBe(true)
  })
  it('picks the volume glyph and mute/volume transitions', () => {
    expect(volumeLevel(1, false)).toBe('high')
    expect(volumeLevel(0.3, false)).toBe('low')
    expect(volumeLevel(1, true)).toBe('mute')
    expect(volumeLevel(0, false)).toBe('mute')
    expect(toggleMute(0.7, false)).toEqual({ volume: 0.7, muted: true })
    expect(toggleMute(0.7, true)).toEqual({ volume: 0.7, muted: false })
    expect(toggleMute(0, true)).toEqual({ volume: 0.5, muted: false })
    expect(applyVolumeInput(0)).toEqual({ volume: 0, muted: true })
    expect(applyVolumeInput(0.4)).toEqual({ volume: 0.4, muted: false })
    expect(applyVolumeInput(3)).toEqual({ volume: 1, muted: false })
  })
})

describe('videoBarLayout', () => {
  it('always fits inside the box and sheds parts as the video shrinks', () => {
    const big = videoBarLayout(640, 360)
    expect(big).toMatchObject({ hidden: false, showTime: true, showVolumeSlider: true })
    expect(big.height).toBeLessThanOrEqual(360)
    const mid = videoBarLayout(300, 200)
    expect(mid).toMatchObject({ showTime: true, showVolume: true, showVolumeSlider: false })
    const small = videoBarLayout(180, 100)
    expect(small).toMatchObject({ showTime: false, showVolume: true })
    expect(videoBarLayout(60, 40).hidden).toBe(true)
  })
})

describe('Space ownership', () => {
  it('lets a video claim Space only', () => {
    let hovering = false
    const off = registerShowVideoSpaceClaimer(() => hovering)
    expect(claimShowSpace(key(' '))).toBe(false)
    hovering = true
    expect(claimShowSpace(key(' '))).toBe(true)
    expect(claimShowSpace(key('ArrowRight'))).toBe(false)
    expect(claimShowSpace(key(' ', { metaKey: true }))).toBe(false)
    off()
    expect(claimShowSpace(key(' '))).toBe(false)
  })
})
