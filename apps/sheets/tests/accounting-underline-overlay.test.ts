import { describe, expect, it } from 'vitest'

import {
  accountingOverlayStrokes,
  overlayAngleOf,
} from '../src/renderer/accounting-underline-overlay'

const base = {
  width: 96,
  height: 20,
  mode: 'singleAccounting' as const,
  angle: 0,
  textLength: 0,
  horizontal: 'left' as const,
  vertical: 'bottom' as const,
}

describe('accounting underline overlay geometry', () => {
  it('draws a near-full-width bottom rule for rich-text cells (Excel Mac E1)', () => {
    const [stroke, ...rest] = accountingOverlayStrokes(base)
    expect(rest).toHaveLength(0)
    expect(stroke!.y1).toBe(stroke!.y2)
    expect(stroke!.x2 - stroke!.x1).toBeGreaterThan(90)
    expect(stroke!.y1).toBeGreaterThan(17)
  })

  it('stacks two rules for double accounting', () => {
    const strokes = accountingOverlayStrokes({ ...base, mode: 'doubleAccounting' })
    expect(strokes).toHaveLength(2)
    expect(strokes[0]!.y1 - strokes[1]!.y1).toBeGreaterThanOrEqual(2)
  })

  it('follows the rotated baseline diagonally for a 45 degree cell (Excel Mac F1)', () => {
    const [stroke] = accountingOverlayStrokes({
      ...base,
      width: 64,
      height: 40,
      angle: 45,
      textLength: 40,
    })
    const dx = stroke!.x2 - stroke!.x1
    const dy = stroke!.y1 - stroke!.y2
    expect(dx).toBeGreaterThan(20)
    expect(dy).toBeCloseTo(dx, 5)
  })

  it('mirrors for negative angles and refuses degenerate boxes', () => {
    const [stroke] = accountingOverlayStrokes({ ...base, angle: -30, textLength: 30 })
    expect(stroke!.y2).toBeGreaterThan(stroke!.y1)
    expect(accountingOverlayStrokes({ ...base, width: 1 })).toEqual([])
  })

  it('reads Univer rotation, ignoring vertical stacking', () => {
    expect(overlayAngleOf({ tr: { a: 45, v: 0 } })).toBe(45)
    expect(overlayAngleOf({ tr: { a: 0, v: 1 } })).toBe(0)
    expect(overlayAngleOf(undefined)).toBe(0)
  })
})
