import { describe, expect, it } from 'vitest'
import { FONT_SIZE_LADDER, snapFontSizePt, stepFontSizePt } from '../src/font-size-step'

describe('stepFontSizePt (PowerPoint Increase / Decrease Font Size)', () => {
  it('walks the standard ladder upward one stop per click', () => {
    let s = 8
    const seen = [s]
    while (s < 96) seen.push((s = stepFontSizePt(s, 1)))
    expect(seen).toEqual([...FONT_SIZE_LADDER])
  })

  it('walks the ladder downward and is the exact inverse of growing', () => {
    for (let i = 1; i < FONT_SIZE_LADDER.length; i++) {
      expect(stepFontSizePt(FONT_SIZE_LADDER[i]!, -1)).toBe(FONT_SIZE_LADDER[i - 1])
    }
  })

  it('12 → 14 → 16 and 12 → 11 → 10.5 → 10 (never 13, never stuck)', () => {
    expect(stepFontSizePt(12, 1)).toBe(14)
    expect(stepFontSizePt(14, 1)).toBe(16)
    expect(stepFontSizePt(12, -1)).toBe(11)
    expect(stepFontSizePt(11, -1)).toBe(10.5)
    expect(stepFontSizePt(10.5, -1)).toBe(10)
  })

  it('off-ladder sizes land on the next stop in the direction pressed', () => {
    expect(stepFontSizePt(13, 1)).toBe(14)
    expect(stepFontSizePt(13, -1)).toBe(12)
    expect(stepFontSizePt(12.5, 1)).toBe(14)
    expect(stepFontSizePt(12.5, -1)).toBe(12)
    expect(stepFontSizePt(45, 1)).toBe(48)
    expect(stepFontSizePt(45, -1)).toBe(44)
  })

  it('a size that is only a rounding error away from a stop counts as that stop', () => {
    // 12pt that went through px (or a 1/100pt sz) and came back as 11.99 / 12.0000001
    for (const noisy of [11.99, 11.9999, 12.0000001, 12.01, 12.2]) {
      expect(stepFontSizePt(noisy, 1)).toBe(14)
      expect(stepFontSizePt(noisy, -1)).toBe(11)
    }
  })

  it('every click changes what the size box displays (half-point resolution)', () => {
    for (let pt = 1; pt <= 400; pt += 0.01) {
      const shown = snapFontSizePt(pt)
      const up = stepFontSizePt(pt, 1)
      const down = stepFontSizePt(pt, -1)
      expect(snapFontSizePt(up)).toBeGreaterThan(shown)
      if (shown > 1) expect(snapFontSizePt(down)).toBeLessThan(shown)
    }
  })

  it('steps whole points below 8, down to 1', () => {
    expect(stepFontSizePt(8, -1)).toBe(7)
    expect(stepFontSizePt(7.5, -1)).toBe(7)
    expect(stepFontSizePt(2, -1)).toBe(1)
    expect(stepFontSizePt(1, -1)).toBe(1)
    expect(stepFontSizePt(7, 1)).toBe(8)
    expect(stepFontSizePt(3, 1)).toBe(4)
  })

  it('continues in tens past 96 and tops out at 4000', () => {
    expect(stepFontSizePt(96, 1)).toBe(100)
    expect(stepFontSizePt(100, 1)).toBe(110)
    expect(stepFontSizePt(110, -1)).toBe(100)
    expect(stepFontSizePt(100, -1)).toBe(96)
    expect(stepFontSizePt(4000, 1)).toBe(4000)
    expect(stepFontSizePt(3995, 1)).toBe(4000)
  })

  it('treats a missing size as the 18pt default', () => {
    expect(stepFontSizePt(Number.NaN, 1)).toBe(20)
  })
})
