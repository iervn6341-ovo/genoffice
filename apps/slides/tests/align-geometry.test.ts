import { describe, expect, it } from 'vitest'
import { alignPositions, minForDistribute, boundingBox } from '../src/renderer/align-geometry'

const page = { w: 1000, h: 500 }
const A = { x: 100, y: 50, w: 100, h: 80 }
const B = { x: 400, y: 200, w: 200, h: 60 }
const C = { x: 800, y: 300, w: 50, h: 100 }

describe('alignPositions: Align Selected Objects (default)', () => {
  it('aligns to the selection bounding box', () => {
    expect(alignPositions([A, B], 'left', 'selection', page)).toEqual([
      { x: 100, y: 50 },
      { x: 100, y: 200 },
    ])
    expect(alignPositions([A, B], 'right', 'selection', page)!.map((p) => p.x)).toEqual([500, 400])
    // bbox y 50..260 -> centre 155
    expect(alignPositions([A, B], 'center-v', 'selection', page)!.map((p) => p.y)).toEqual([
      115, 125,
    ])
  })

  it('a single element aligns to the slide even in selection mode', () => {
    expect(alignPositions([A], 'left', 'selection', page)).toEqual([{ x: 0, y: 50 }])
    expect(alignPositions([A], 'bottom', 'selection', page)).toEqual([{ x: 100, y: 420 }])
  })

  it('distribute keeps the outer two and evens the gaps; needs three', () => {
    expect(alignPositions([A, B], 'distribute-h', 'selection', page)).toBeNull()
    const out = alignPositions([A, B, C], 'distribute-h', 'selection', page)!
    // span 100..850 = 750, widths 350 -> gap 200
    expect(out.map((p) => p.x)).toEqual([100, 400, 800])
    const moved = alignPositions([A, { ...B, x: 300 }, C], 'distribute-h', 'selection', page)!
    expect(moved[1]!.x).toBe(400)
  })
})

describe('alignPositions: Align to Slide', () => {
  it('uses the slide edges for every element, not their bounding box', () => {
    expect(alignPositions([A, B], 'left', 'slide', page)!.map((p) => p.x)).toEqual([0, 0])
    expect(alignPositions([A, B], 'right', 'slide', page)!.map((p) => p.x)).toEqual([900, 800])
    expect(alignPositions([A, B], 'center-h', 'slide', page)!.map((p) => p.x)).toEqual([450, 400])
    expect(alignPositions([A, B], 'bottom', 'slide', page)!.map((p) => p.y)).toEqual([420, 440])
  })

  it('distribute spreads first..last across the slide and works with two elements', () => {
    expect(minForDistribute('slide')).toBe(2)
    expect(minForDistribute('selection')).toBe(3)
    const two = alignPositions([A, B], 'distribute-h', 'slide', page)!
    // widths 300 on a 1000 slide: first at 0, last flush right
    expect(two.map((p) => p.x)).toEqual([0, 800])
    const three = alignPositions([C, A, B], 'distribute-h', 'slide', page)!
    // order by x: A(100) B(200) C(50): gap (1000-350)/2 = 325
    expect(three[1]!.x).toBe(0)
    expect(three[2]!.x).toBe(425)
    expect(three[0]!.x).toBe(950)
    // y is untouched by a horizontal distribute
    expect(three.map((p) => p.y)).toEqual([300, 50, 200])
  })

  it('vertical distribute uses the slide height', () => {
    const out = alignPositions([A, B], 'distribute-v', 'slide', page)!
    expect(out.map((p) => p.y)).toEqual([0, 440])
  })

  it('a lone element cannot be distributed', () => {
    expect(alignPositions([A], 'distribute-h', 'slide', page)).toBeNull()
  })
})

describe('boundingBox', () => {
  it('spans all rects', () => {
    expect(boundingBox([A, C])).toEqual({ x: 100, y: 50, w: 750, h: 350 })
  })
})
