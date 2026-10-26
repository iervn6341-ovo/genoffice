/**
 * Excel's Increase / Decrease Font Size (Home > Font).
 *
 * The button walks Excel's size ladder (…16, 18, 20, 22, 24, 26, 28, 36, 48, 72…) and does it PER
 * CELL: a selection of a 10 pt and a 20 pt cell becomes 11 pt and 22 pt, not 11 pt and 11 pt. Past
 * the ends the ladder continues the way Word's does - +10 above 72 (capped at Excel's 409) and -1
 * below 8 (floor 1) - so the button never goes dead.
 */

export const FONT_SIZE_LADDER: readonly number[] = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
]

/** Excel's font-size bounds. */
export const MIN_FONT_SIZE = 1
export const MAX_FONT_SIZE = 409

/** Past this many cells a step is applied from the first cell's size to the whole selection. */
export const PER_CELL_STEP_LIMIT = 20_000

export function stepFontSize(current: number, direction: 1 | -1): number {
  const top = FONT_SIZE_LADDER[FONT_SIZE_LADDER.length - 1] ?? 72
  const bottom = FONT_SIZE_LADDER[0] ?? 8
  if (!Number.isFinite(current) || current <= 0) current = 11
  if (direction === 1) {
    if (current >= top) return Math.min(MAX_FONT_SIZE, Math.floor(current / 10) * 10 + 10)
    return FONT_SIZE_LADDER.find((size) => size > current) ?? top
  }
  if (current > top) {
    const lower = Math.ceil(current / 10) * 10 - 10
    return lower <= top ? top : lower
  }
  const lower = [...FONT_SIZE_LADDER].reverse().find((size) => size < current)
  if (lower !== undefined) return lower
  return Math.max(MIN_FONT_SIZE, Math.ceil(current) - 1)
}

export interface CellBlock {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface SizedBlock extends CellBlock {
  readonly size: number
}

/**
 * The rectangles to restyle for one Increase / Decrease press over `ranges`, each cell stepping
 * from its own size. Cells that appear in several ranges step once. Neighbouring cells that land
 * on the same size are merged (runs along a row, then identical runs on consecutive rows) so a
 * uniform selection is one rectangle and a mixed one a handful.
 */
export function planFontSizeStep(
  ranges: readonly CellBlock[],
  sizeAt: (row: number, column: number) => number,
  direction: 1 | -1,
): SizedBlock[] {
  const seen = new Set<string>()
  const rows = new Map<number, Array<{ column: number; size: number }>>()
  for (const range of ranges) {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const key = `${row},${column}`
        if (seen.has(key)) continue
        seen.add(key)
        const list = rows.get(row) ?? []
        list.push({ column, size: stepFontSize(sizeAt(row, column), direction) })
        rows.set(row, list)
      }
    }
  }
  type Run = { startColumn: number; endColumn: number; size: number }
  const runsByRow = new Map<number, Run[]>()
  for (const [row, cells] of rows) {
    cells.sort((a, b) => a.column - b.column)
    const runs: Run[] = []
    for (const cell of cells) {
      const last = runs[runs.length - 1]
      if (last && last.size === cell.size && last.endColumn === cell.column - 1)
        last.endColumn = cell.column
      else runs.push({ startColumn: cell.column, endColumn: cell.column, size: cell.size })
    }
    runsByRow.set(row, runs)
  }
  const blocks: Array<SizedBlock & { endRow: number }> = []
  const open = new Map<string, number>() // run signature -> index in blocks of the block still growing
  let previousRow = Number.NaN
  for (const row of [...runsByRow.keys()].sort((a, b) => a - b)) {
    const nextOpen = new Map<string, number>()
    for (const run of runsByRow.get(row) ?? []) {
      const signature = `${run.startColumn}:${run.endColumn}:${run.size}`
      const grow = row === previousRow + 1 ? open.get(signature) : undefined
      if (grow !== undefined) {
        const block = blocks[grow]!
        blocks[grow] = { ...block, endRow: row }
        nextOpen.set(signature, grow)
      } else {
        blocks.push({
          startRow: row,
          endRow: row,
          startColumn: run.startColumn,
          endColumn: run.endColumn,
          size: run.size,
        })
        nextOpen.set(signature, blocks.length - 1)
      }
    }
    open.clear()
    for (const [key, value] of nextOpen) open.set(key, value)
    previousRow = row
  }
  return blocks
}
