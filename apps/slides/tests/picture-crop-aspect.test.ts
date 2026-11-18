import { describe, expect, it } from 'vitest'
import {
  computeAspectRatioCrop,
  computeFillCrop,
  computeFitCrop,
  reconstructFullBox,
} from '../src/renderer/picture-edit-actions'

describe('reconstructFullBox', () => {
  it('returns box directly when uncropped', () => {
    const box = { x: 100, y: 50, w: 200, h: 100 }
    expect(reconstructFullBox(box, null)).toEqual(box)
  })

  it('reconstructs original extent when cropped', () => {
    // 50% horizontal crop (l=0.25, r=0.25)
    const box = { x: 150, y: 50, w: 100, h: 100 }
    const full = reconstructFullBox(box, { l: 0.25, r: 0.25, t: 0, b: 0 })
    expect(full.w).toBeCloseTo(200)
    expect(full.x).toBeCloseTo(100)
    expect(full.h).toBeCloseTo(100)
    expect(full.y).toBeCloseTo(50)
  })
})

describe('computeAspectRatioCrop', () => {
  it('applies 1:1 square crop to a 2:1 landscape image', () => {
    const fullBox = { x: 0, y: 0, w: 200, h: 100 }
    const box = { x: 0, y: 0, w: 200, h: 100 }
    const res = computeAspectRatioCrop(box, fullBox, 1, 1)

    // Horizontal crop should be 25% from left and 25% from right (leaving center 50%)
    expect(res.srcRect.l).toBeCloseTo(0.25)
    expect(res.srcRect.r).toBeCloseTo(0.25)
    expect(res.srcRect.t).toBe(0)
    expect(res.srcRect.b).toBe(0)

    // The resulting box should be square and centered
    expect(res.boxPx.w).toBeCloseTo(res.boxPx.h)
    expect(res.boxPx.h).toBeCloseTo(100)
    expect(res.boxPx.w).toBeCloseTo(100)
    expect(res.boxPx.x).toBeCloseTo(50)
    expect(res.boxPx.y).toBeCloseTo(0)
  })

  it('applies 16:9 crop to a 1:1 square image', () => {
    const fullBox = { x: 0, y: 0, w: 100, h: 100 }
    const box = { x: 0, y: 0, w: 100, h: 100 }
    const res = computeAspectRatioCrop(box, fullBox, 16, 9)

    // Vertical crop to 9/16 = 0.5625, so crop 0.21875 from top and bottom
    expect(res.srcRect.l).toBe(0)
    expect(res.srcRect.r).toBe(0)
    expect(res.srcRect.t).toBeCloseTo((1 - 9 / 16) / 2)
    expect(res.srcRect.b).toBeCloseTo((1 - 9 / 16) / 2)

    expect(res.boxPx.w).toBeCloseTo(100)
    expect(res.boxPx.h).toBeCloseTo((100 * 9) / 16)
  })
})

describe('computeFillCrop', () => {
  it('crops sides to fill a squarer frame with a wide image', () => {
    const fullBox = { x: 0, y: 0, w: 200, h: 100 }
    const box = { x: 10, y: 10, w: 100, h: 100 }
    const res = computeFillCrop(box, fullBox)

    expect(res.srcRect).not.toBeNull()
    expect(res.srcRect!.l).toBeCloseTo(0.25)
    expect(res.srcRect!.r).toBeCloseTo(0.25)
    expect(res.boxPx).toEqual(box)
  })
})

describe('computeFitCrop', () => {
  it('fits wide image into square box by adjusting height and centering', () => {
    const fullBox = { x: 0, y: 0, w: 200, h: 100 }
    const box = { x: 0, y: 0, w: 100, h: 100 }
    const res = computeFitCrop(box, fullBox)

    expect(res.srcRect).toBeNull()
    expect(res.boxPx.w).toBeCloseTo(100)
    expect(res.boxPx.h).toBeCloseTo(50)
    expect(res.boxPx.x).toBeCloseTo(0)
    expect(res.boxPx.y).toBeCloseTo(25)
  })
})
