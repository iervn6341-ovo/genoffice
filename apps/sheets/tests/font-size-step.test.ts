import { describe, expect, it } from 'vitest'

import { FONT_SIZE_LADDER, planFontSizeStep, stepFontSize } from '../src/renderer/font-size-step'

describe('stepFontSize', () => {
  it('has 22 between 20 and 24 (Excel ladder)', () => {
    expect(FONT_SIZE_LADDER).toContain(22)
    expect(stepFontSize(20, 1)).toBe(22)
    expect(stepFontSize(22, 1)).toBe(24)
    expect(stepFontSize(24, -1)).toBe(22)
  })

  it('climbs the full ladder from 11', () => {
    const seen: number[] = []
    let size = 11
    for (let i = 0; i < 12; i += 1) {
      size = stepFontSize(size, 1)
      seen.push(size)
    }
    expect(seen).toEqual([12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72])
  })

  it('walks down 11 -> 8 then one point at a time to 1', () => {
    expect(stepFontSize(11, -1)).toBe(10)
    expect(stepFontSize(9, -1)).toBe(8)
    expect(stepFontSize(8, -1)).toBe(7)
    expect(stepFontSize(2, -1)).toBe(1)
    expect(stepFontSize(1, -1)).toBe(1)
  })

  it('continues past 72 by tens and caps at 409', () => {
    expect(stepFontSize(72, 1)).toBe(80)
    expect(stepFontSize(80, 1)).toBe(90)
    expect(stepFontSize(405, 1)).toBe(409)
    expect(stepFontSize(409, 1)).toBe(409)
    expect(stepFontSize(90, -1)).toBe(80)
    expect(stepFontSize(80, -1)).toBe(72)
  })

  it('snaps a non-ladder size to the next rung', () => {
    expect(stepFontSize(10.5, 1)).toBe(11)
    expect(stepFontSize(10.5, -1)).toBe(10)
    expect(stepFontSize(13, 1)).toBe(14)
    expect(stepFontSize(13, -1)).toBe(12)
  })
})

describe('planFontSizeStep', () => {
  const range = (r1: number, c1: number, r2 = r1, c2 = c1) => ({
    startRow: r1,
    endRow: r2,
    startColumn: c1,
    endColumn: c2,
  })

  it('steps every cell from its own size (mixed sizes stay distinct)', () => {
    const sizes: Record<string, number> = { '0,0': 10, '1,0': 20 }
    const plan = planFontSizeStep([range(0, 0, 1, 0)], (r, c) => sizes[`${r},${c}`] ?? 11, 1)
    expect(plan).toEqual([
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0, size: 11 },
      { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0, size: 22 },
    ])
  })

  it('merges a uniform selection into one rectangle', () => {
    expect(planFontSizeStep([range(0, 0, 9, 3)], () => 11, 1)).toEqual([
      { startRow: 0, endRow: 9, startColumn: 0, endColumn: 3, size: 12 },
    ])
  })

  it('reaches every range of a multi-range selection and steps overlaps once', () => {
    const plan = planFontSizeStep(
      [range(0, 0, 2, 0), range(0, 2, 2, 2), range(2, 0, 2, 2)],
      () => 11,
      1,
    )
    const covered = new Set<string>()
    for (const b of plan)
      for (let r = b.startRow; r <= b.endRow; r += 1)
        for (let c = b.startColumn; c <= b.endColumn; c += 1) {
          expect(covered.has(`${r},${c}`)).toBe(false)
          covered.add(`${r},${c}`)
          expect(b.size).toBe(12)
        }
    expect(covered.size).toBe(7)
    expect(covered.has('0,2')).toBe(true)
  })

  it('does not merge non-adjacent cells of equal size', () => {
    const plan = planFontSizeStep([range(0, 0), range(0, 2)], () => 11, -1)
    expect(plan).toHaveLength(2)
  })
})
