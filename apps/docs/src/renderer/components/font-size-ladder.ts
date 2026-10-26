/**
 * Word's Grow Font / Shrink Font ladder (measured with Word for Mac's `grow font` / `shrink font`):
 *   1,2,3,4 (one point apart) · 5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24,
 *   26, 28, 36, 48, 72 · 80, 90, 100 … 150 (ten apart).
 * A value off the ladder (11.5, 75) jumps to the nearest rung in the stepping direction. Beyond 150
 * the ten-point step continues up to Word's 1638pt ceiling (inferred — the measured run stopped at 150).
 */
export const SIZE_LADDER_MIN = 1
export const SIZE_LADDER_MAX = 1638

const MID_LADDER = [
  5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
]

export function growFontSize(size: number): number {
  if (!(size > 0)) return 5
  if (size < 5) return Math.floor(size) + 1
  if (size >= 72) return Math.min(Math.floor(size / 10) * 10 + 10, SIZE_LADDER_MAX)
  return MID_LADDER.find((s) => s > size) ?? 72
}

export function shrinkFontSize(size: number): number {
  if (!(size > 0)) return SIZE_LADDER_MIN
  if (size <= 5) return Math.max(Math.ceil(size) - 1, SIZE_LADDER_MIN)
  if (size > 72) return Math.max(Math.ceil(size / 10) * 10 - 10, 72)
  for (let i = MID_LADDER.length - 1; i >= 0; i--) if (MID_LADDER[i] < size) return MID_LADDER[i]
  return 5
}

export const stepFontSizeLadder = (size: number, dir: 1 | -1): number =>
  dir === 1 ? growFontSize(size) : shrinkFontSize(size)
