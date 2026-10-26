import { describe, expect, it } from 'vitest'
import { placeContextMenu } from '../src/renderer/context-menu-position'

describe('placeContextMenu', () => {
  it('leaves a menu that fits where it was opened', () => {
    expect(placeContextMenu(100, 100, 200, 300, 1280, 760)).toEqual({
      x: 100,
      y: 100,
      maxHeight: null,
    })
  })

  it('slides a menu up just enough to clear the bottom edge', () => {
    // 198px thumbnail menu opened 20px above the bottom of a 760px window: the old 190px guess clipped 8px
    const p = placeContextMenu(50, 740, 180, 198, 1280, 760)
    expect(p.maxHeight).toBeNull()
    expect(p.y + 198).toBeLessThanOrEqual(760 - 4)
  })

  it('keeps the menu inside on the right edge', () => {
    const p = placeContextMenu(1270, 50, 200, 100, 1280, 760)
    expect(p.x + 200).toBeLessThanOrEqual(1280 - 4)
  })

  it('caps a menu taller than the window (29-row table menu, 887px in 760px) so it can scroll', () => {
    const p = placeContextMenu(400, 300, 220, 887, 1280, 760)
    expect(p.maxHeight).toBe(752)
    expect(p.y).toBe(4)
    expect(p.y + p.maxHeight!).toBeLessThanOrEqual(760)
  })

  it('never returns a negative position or cap in a tiny window', () => {
    const p = placeContextMenu(10, 10, 300, 500, 200, 6)
    expect(p.x).toBeGreaterThanOrEqual(0)
    expect(p.maxHeight).toBe(0)
  })
})
