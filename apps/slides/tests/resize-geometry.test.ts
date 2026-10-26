import { describe, expect, it } from 'vitest'
import { centerShift, contentBoundHeight, drawnTableHeight } from '../src/renderer/resize-geometry'

describe('centerShift', () => {
  const from = { w: 200, h: 100 }
  it('a bottom handle pins the top edge: the centre follows the growing bottom', () => {
    expect(centerShift('bottom-right', 0, from, { w: 200, h: 140 })).toEqual({ dx: 0, dy: 20 })
  })
  it('a top handle pins the bottom edge', () => {
    expect(centerShift('top-left', 0, from, { w: 200, h: 140 })).toEqual({ dx: 0, dy: -20 })
  })
  it('left handles pin the right edge, right handles the left edge', () => {
    expect(centerShift('middle-left', 0, from, { w: 260, h: 100 }).dx).toBe(-30)
    expect(centerShift('middle-right', 0, from, { w: 260, h: 100 }).dx).toBe(30)
  })
  it('a middle handle leaves the other axis alone', () => {
    expect(centerShift('middle-right', 0, from, { w: 200, h: 180 }).dy).toBe(0)
    expect(centerShift('top-center', 0, from, { w: 400, h: 100 }).dx).toBe(0)
  })
  it('rotates the shift into stage space', () => {
    const r = centerShift('bottom-center', 90, from, { w: 200, h: 140 })
    expect(r.dx).toBeCloseTo(-20, 6)
    expect(r.dy).toBeCloseTo(0, 6)
  })
  it('no anchor, no shift', () => {
    expect(centerShift(undefined, 0, from, { w: 10, h: 10 })).toEqual({ dx: 0, dy: 0 })
  })
})

describe('table drawn height', () => {
  it('follows the content when the layout is taller than the pointer asked', () => {
    expect(drawnTableHeight(120, 255)).toBe(255)
    expect(drawnTableHeight(300, 255)).toBe(300)
    expect(drawnTableHeight(120, 0)).toBe(120)
  })
  it('a reply is content-bound only when it exceeds what was asked for', () => {
    expect(contentBoundHeight(120, 255)).toBe(255)
    expect(contentBoundHeight(120, 120.3)).toBe(0)
    expect(contentBoundHeight(null, 255)).toBe(0)
    expect(contentBoundHeight(300, 255)).toBe(0)
  })
})
