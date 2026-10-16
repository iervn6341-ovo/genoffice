import { describe, expect, it } from 'vitest'
import { showSetupOrder } from '../src/renderer/slideshow-utils'

describe('native show selection playback', () => {
  it('maps one-based inclusive range to original slide indices', () => {
    expect(
      showSetupOrder(
        {
          mode: 'browse',
          loop: false,
          showNarration: true,
          showAnimation: true,
          useTimings: false,
          slides: { kind: 'range', start: 2, end: 4 },
        },
        5,
        [],
      ),
    ).toEqual([1, 2, 3])
  })

  it('uses native custom show order and retains repeated slide references', () => {
    expect(
      showSetupOrder(
        {
          mode: 'speaker',
          loop: true,
          showNarration: false,
          showAnimation: false,
          useTimings: true,
          slides: { kind: 'custom', id: '7' },
        },
        5,
        [{ id: '7', name: 'Review', slideIndices: [3, 1, 3, 9] }],
      ),
    ).toEqual([3, 1, 3])
  })
})
