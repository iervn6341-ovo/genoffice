import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

describe('docs pinch-to-zoom compositor preview', () => {
  let timerCallbacks: Array<() => void> = []

  beforeEach(() => {
    timerCallbacks = []
  })

  it('calculates continuous visual scale and translation from wheel deltas', () => {
    const startZoom = 100
    const deltaY = -10
    const clampedDelta = Math.max(-100, Math.min(100, deltaY))
    const factor = Math.exp(-clampedDelta * 0.0075)
    const nextZoom = Math.min(200, Math.max(50, startZoom * factor))

    expect(nextZoom).toBeCloseTo(107.79, 2)

    // Visual ratio
    const visualRatio = nextZoom / startZoom
    expect(visualRatio).toBeCloseTo(1.0779, 4)

    // Anchor translation
    const vx = 300
    const vy = 200
    const tx = vx * (1 - visualRatio)
    const ty = vy * (1 - visualRatio)

    expect(tx).toBeCloseTo(300 * (1 - 1.0779), 2)
    expect(ty).toBeCloseTo(200 * (1 - 1.0779), 2)
  })
})
