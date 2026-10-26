import { describe, expect, it } from 'vitest'
import { clampMenuToViewport } from '../src/renderer/menu-clamp'

describe('clampMenuToViewport', () => {
  const size = { width: 170, height: 198 }
  it('leaves a menu that fits where the pointer is', () => {
    expect(clampMenuToViewport({ x: 100, y: 100 }, size, { width: 1000, height: 760 })).toEqual({
      x: 100,
      y: 100,
    })
  })
  it('uses the measured height, not a guess (198px menu at the bottom of a 760px window)', () => {
    const r = clampMenuToViewport({ x: 100, y: 740 }, size, { width: 1000, height: 760 })
    expect(r.y + size.height).toBeLessThanOrEqual(760 - 4)
    expect(r.y).toBe(558)
  })
  it('clamps at the right edge and never goes off the top/left', () => {
    expect(clampMenuToViewport({ x: 990, y: 10 }, size, { width: 1000, height: 760 }).x).toBe(826)
    expect(clampMenuToViewport({ x: -5, y: -5 }, size, { width: 1000, height: 760 })).toEqual({
      x: 4,
      y: 4,
    })
  })
})
