import { describe, expect, it } from 'vitest'

import { DEFAULT_AI_PANEL_PREFS, normalizeAiPanelPrefs } from '../src/ai-panel-prefs'
import {
  SMOOTH_STREAM_TAU_MS,
  SMOOTH_TAU_MS,
  STREAM_GAP_MS,
  easeStep,
  isNotchedWheel,
  isSmoothable,
  tauForGap,
  wheelPixels,
} from '../src/smooth-scroll'

const wheel = (over: Partial<Parameters<typeof isNotchedWheel>[0]>) => ({
  deltaX: 0,
  deltaY: 0,
  deltaMode: 0,
  ctrlKey: false,
  metaKey: false,
  wheelDeltaY: 0,
  ...over,
})

describe('isNotchedWheel', () => {
  it('takes over a Windows-style notched wheel (100 px = 120 wheel units)', () => {
    expect(isNotchedWheel(wheel({ deltaY: 100, wheelDeltaY: -120 }))).toBe(true)
    expect(isNotchedWheel(wheel({ deltaY: 200, wheelDeltaY: -240 }))).toBe(true)
  })
  it('takes over line-based wheels', () => {
    expect(isNotchedWheel(wheel({ deltaY: 3, deltaMode: 1 }))).toBe(true)
  })
  it('leaves trackpad and momentum deltas alone', () => {
    expect(isNotchedWheel(wheel({ deltaY: 3.5, wheelDeltaY: -10.5 }))).toBe(false)
    expect(isNotchedWheel(wheel({ deltaY: 12, wheelDeltaY: -36 }))).toBe(false)
  })
  it('never touches zoom gestures', () => {
    expect(isNotchedWheel(wheel({ deltaY: 100, wheelDeltaY: -120, ctrlKey: true }))).toBe(false)
    expect(isNotchedWheel(wheel({ deltaY: 100, wheelDeltaY: -120, metaKey: true }))).toBe(false)
  })
})

describe('what gets smoothed and how fast', () => {
  it('smooths notches, trackpad deltas and line wheels, but never zoom gestures or zero deltas', () => {
    expect(isSmoothable(wheel({ deltaY: 100, wheelDeltaY: -120 }))).toBe(true)
    expect(isSmoothable(wheel({ deltaY: 3.5 }))).toBe(true)
    expect(isSmoothable(wheel({ deltaX: 12 }))).toBe(true)
    expect(isSmoothable(wheel({ deltaY: 100, ctrlKey: true }))).toBe(false)
    expect(isSmoothable(wheel({ deltaY: 100, metaKey: true }))).toBe(false)
    expect(isSmoothable(wheel({}))).toBe(false)
  })
  it('uses a longer glide for isolated notches and a short one inside a fast stream', () => {
    expect(tauForGap(120)).toBe(SMOOTH_TAU_MS)
    expect(tauForGap(STREAM_GAP_MS)).toBe(SMOOTH_TAU_MS)
    expect(tauForGap(STREAM_GAP_MS - 1)).toBe(SMOOTH_STREAM_TAU_MS)
    expect(tauForGap(16)).toBeLessThan(tauForGap(200))
    expect(tauForGap(Infinity)).toBe(SMOOTH_TAU_MS)
  })
  it('a shorter time constant catches up faster', () => {
    expect(easeStep(100, 16, SMOOTH_STREAM_TAU_MS)).toBeGreaterThan(
      easeStep(100, 16, SMOOTH_TAU_MS),
    )
  })
})

describe('wheelPixels / easeStep', () => {
  it('converts line and page deltas to pixels', () => {
    expect(wheelPixels(3, 1, 800)).toBe(120)
    expect(wheelPixels(1, 2, 800)).toBe(800)
    expect(wheelPixels(100, 0, 800)).toBe(100)
  })
  it('eases toward the target without overshooting, and lands exactly at the end', () => {
    let remaining = 300
    let moved = 0
    for (let frame = 0; frame < 120 && remaining !== 0; frame += 1) {
      const step = easeStep(remaining, 16)
      expect(step).toBeGreaterThan(0)
      expect(step).toBeLessThanOrEqual(remaining)
      remaining -= step
      moved += step
    }
    expect(remaining).toBe(0)
    expect(moved).toBeCloseTo(300, 5)
  })
  it('covers ~95% of a notch in three time constants and scales with frame time', () => {
    expect(easeStep(100, SMOOTH_TAU_MS * 3)).toBeGreaterThan(94)
    expect(easeStep(100, 32)).toBeGreaterThan(easeStep(100, 16))
    expect(easeStep(-100, 16)).toBeLessThan(0)
    expect(easeStep(100, 0)).toBe(0)
  })
})

describe('display prefs', () => {
  it('defaults: smooth scrolling off, GPU on; malformed values fall back', () => {
    expect(DEFAULT_AI_PANEL_PREFS.smoothScroll).toBe(false)
    expect(DEFAULT_AI_PANEL_PREFS.gpuAcceleration).toBe(true)
    const prefs = normalizeAiPanelPrefs({ smoothScroll: 'yes', gpuAcceleration: 0 })
    expect(prefs.smoothScroll).toBe(false)
    expect(prefs.gpuAcceleration).toBe(true)
    expect(normalizeAiPanelPrefs({ smoothScroll: true, gpuAcceleration: false })).toMatchObject({
      smoothScroll: true,
      gpuAcceleration: false,
    })
  })
})
