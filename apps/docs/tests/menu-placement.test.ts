import { describe, expect, it } from 'vitest'
import { placeContextMenu, placeSubmenu } from '../src/renderer/components/menu-placement'

const vp = { width: 1000, height: 700 }

describe('placeContextMenu', () => {
  it('opens down-right at the pointer when it fits', () => {
    expect(placeContextMenu({ x: 100, y: 100 }, { width: 240, height: 300 }, vp)).toEqual({
      left: 100,
      top: 100,
    })
  })
  it('flips up near the bottom and clamps near the right edge', () => {
    expect(placeContextMenu({ x: 900, y: 600 }, { width: 240, height: 300 }, vp)).toEqual({
      left: 752,
      top: 300,
    })
  })
  it('pins to the bottom when it neither fits below nor above', () => {
    // 450px menu, pointer at y=300 in a 700px window: below runs out, above has only 300
    expect(placeContextMenu({ x: 10, y: 300 }, { width: 240, height: 450 }, vp)).toEqual({
      left: 10,
      top: 242,
    })
  })
  it('caps a menu taller than the window and lets it scroll (450px-tall window)', () => {
    const r = placeContextMenu(
      { x: 50, y: 200 },
      { width: 240, height: 468 },
      { width: 800, height: 450 },
    )
    expect(r).toEqual({ left: 50, top: 8, maxHeight: 434 })
    expect(r.top + r.maxHeight!).toBeLessThanOrEqual(450 - 8)
  })
})

describe('placeSubmenu', () => {
  const anchor = { left: 300, right: 540, top: 200 }
  it('opens to the right, slightly lifted', () => {
    expect(placeSubmenu(anchor, { width: 160, height: 120 }, vp)).toEqual({ left: 536, top: 195 })
  })
  it('flips to the left of the parent when the right side overflows (700px-wide window)', () => {
    const r = placeSubmenu(
      { left: 420, right: 660, top: 200 },
      { width: 160, height: 120 },
      { width: 700, height: 700 },
    )
    expect(r.left).toBe(264)
    expect(r.left + 160).toBeLessThanOrEqual(700 - 8)
  })
  it('stays inside when neither side fits', () => {
    const r = placeSubmenu(
      { left: 10, right: 250, top: 100 },
      { width: 400, height: 100 },
      { width: 300, height: 500 },
    )
    expect(r.left).toBeGreaterThanOrEqual(8)
  })
  it('slides up at the bottom and caps + scrolls a tall submenu', () => {
    expect(
      placeSubmenu({ left: 0, right: 100, top: 650 }, { width: 100, height: 120 }, vp),
    ).toEqual({
      left: 96,
      top: 572,
    })
    expect(
      placeSubmenu({ left: 0, right: 100, top: 300 }, { width: 100, height: 900 }, vp),
    ).toEqual({ left: 96, top: 8, maxHeight: 684 })
  })
})
