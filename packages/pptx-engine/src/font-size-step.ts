/**
 * PowerPoint's Increase / Decrease Font Size ladder (Home → Font, ⌘⇧> / ⌘⇧<). Shared by the
 * editor (selection while typing) and the element-level font patch (selected shapes), so both
 * step to the same sizes.
 */
export const FONT_SIZE_LADDER: readonly number[] = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
]

/** Next/previous ladder size; beyond the ladder ±10pt, clamped to 8~400 */
export function stepFontSizePt(cur: number, dir: 1 | -1): number {
  const ladder = FONT_SIZE_LADDER
  const max = ladder[ladder.length - 1]!
  if (dir > 0) return cur >= max ? Math.min(400, cur + 10) : ladder.find((s) => s > cur)!
  if (cur > max) return Math.max(max, cur - 10)
  for (let i = ladder.length - 1; i >= 0; i--) if (ladder[i]! < cur) return ladder[i]!
  return ladder[0]!
}
