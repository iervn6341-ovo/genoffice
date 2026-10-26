/**
 * Content-aware column floors of a table. A resize gesture (and a typed width) may shrink a column
 * only until its widest unbreakable word still fits (`colNeedPx`, measured with the text layout's own
 * tokenizer), never below a tiny absolute floor, and never above the width the column already has
 * (a column that is already too narrow is not forced wider, it just cannot shrink further).
 */

/** Absolute floor of an empty / short column (px) */
export const TABLE_COL_FLOOR_PX = 24
/** Float slack added to a measured word width so rounding never lets the word tip over the edge (px) */
export const TABLE_COL_FIT_SLACK_PX = 0.5

/** Per-column floor (px): max(content need + slack, `base`), capped at the column's current width. */
export function tableColFloors(
  colWidths: number[],
  colNeeds?: number[],
  base: number = TABLE_COL_FLOOR_PX,
): number[] {
  return colWidths.map((w, i) => {
    const need = colNeeds?.[i] ?? 0
    const want = Math.max(need > 0 ? need + TABLE_COL_FIT_SLACK_PX : 0, base)
    return Math.min(want, Math.max(w, 1))
  })
}

/**
 * Smallest uniform factor the column set can be scaled by (the engine resizes a table by scaling its
 * columns proportionally) before some column drops under its floor. 1 = cannot shrink at all.
 */
export function tableMinScale(colWidths: number[], colNeeds?: number[], base?: number): number {
  const floors = tableColFloors(colWidths, colNeeds, base)
  let k = 0
  colWidths.forEach((w, i) => {
    if (w > 0) k = Math.max(k, floors[i]! / w)
  })
  return Math.min(k, 1)
}
