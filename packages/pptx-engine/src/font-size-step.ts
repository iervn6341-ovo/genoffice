/**
 * PowerPoint's Increase / Decrease Font Size ladder (Home → Font, ⌘⇧> / ⌘⇧<). Shared by the
 * editor (selection while typing) and the element-level font patch (selected shapes and table
 * cells), so both step to the same sizes.
 *
 *   1…8 in whole points → 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48,
 *   54, 60, 66, 72, 80, 88, 96 → then +10 per click up to PowerPoint's 4000 pt ceiling.
 *
 * Every click must visibly change the number the size box shows (it displays half-point
 * resolution), so the current size is first snapped to that same resolution: a stored 11.99 or a
 * px round-trip that landed on 12.0000001 is "12", never "already past 12".
 */
export const FONT_SIZE_LADDER: readonly number[] = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
]

export const FONT_SIZE_MIN_PT = 1
export const FONT_SIZE_MAX_PT = 4000

/** The resolution the ribbon size box displays (and the editor stores): half a point */
export const snapFontSizePt = (pt: number): number => Math.round(pt * 2) / 2

/** Next/previous ladder size from `cur`; always a different value unless clamped at 1 / 4000 */
export function stepFontSizePt(cur: number, dir: 1 | -1): number {
  const c = Number.isFinite(cur) ? snapFontSizePt(cur) : 18
  const ladder = FONT_SIZE_LADDER
  const lo = ladder[0]!
  const hi = ladder[ladder.length - 1]!
  if (dir > 0) {
    if (c < lo) return Math.min(lo, Math.floor(c) + 1) // 1…7: one point at a time up to 8
    if (c >= hi) return Math.min(FONT_SIZE_MAX_PT, Math.floor(c / 10) * 10 + 10)
    return ladder.find((s) => s > c)!
  }
  if (c > hi) return Math.max(hi, Math.ceil(c / 10) * 10 - 10)
  if (c <= lo) return Math.max(FONT_SIZE_MIN_PT, Math.ceil(c) - 1) // 8 → 7 → … → 1
  for (let i = ladder.length - 1; i >= 0; i--) if (ladder[i]! < c) return ladder[i]!
  return lo
}
