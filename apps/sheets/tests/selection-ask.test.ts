import { describe, expect, it } from 'vitest'
import {
  FILL_HANDLE_CLEARANCE_PX,
  isSelectionDrag,
  selectionAskPosition,
} from '../src/renderer/ai/selection-ask'

const viewport = { left: 100, top: 50, right: 900, bottom: 650 }

describe('selection Ask AI gesture', () => {
  it('ignores click jitter and accepts a deliberate drag', () => {
    expect(isSelectionDrag({ x: 20, y: 20 }, { x: 22, y: 21 })).toBe(false)
    expect(isSelectionDrag({ x: 20, y: 20 }, { x: 24, y: 20 })).toBe(true)
  })

  it('places the trigger above-right of the drag endpoint when there is room', () => {
    expect(selectionAskPosition({ x: 300, y: 200 }, viewport, 92, 32)).toEqual({
      left: 308,
      top: 160,
    })
  })

  it('never covers the fill handle below the endpoint', () => {
    const pointer = { x: 300, y: 200 }
    const trigger = selectionAskPosition(pointer, viewport, 92, 32)
    // the handle is within one row (<= 24px) below and one column to the right of the endpoint
    expect(trigger.top + 32).toBeLessThanOrEqual(pointer.y)
    // no room above: drop below the handle's row, not onto it
    const nearTop = selectionAskPosition({ x: 300, y: 60 }, viewport, 92, 32)
    expect(nearTop.top).toBeGreaterThanOrEqual(60 + FILL_HANDLE_CLEARANCE_PX)
  })

  it('flips and clamps the trigger at the grid edges', () => {
    expect(selectionAskPosition({ x: 895, y: 645 }, viewport, 92, 32)).toEqual({
      left: 795,
      top: 605,
    })
  })

  it('uses the measured localized-label width when clamping', () => {
    const position = selectionAskPosition({ x: 895, y: 200 }, viewport, 220, 32)
    expect(position.left).toBe(667)
    expect(position.left + 220).toBeLessThanOrEqual(viewport.right - 4)
  })
})
