/**
 * Excel copies several selected areas as one block when they line up: the same rows (the areas
 * are laid side by side) or the same columns (laid one under the other). Anything else is
 * refused ("That command cannot be used on multiple selections").
 */
import type { IRange } from '@univerjs/core'

export interface MultiSelectionLayout {
  /// the smallest range holding every area
  readonly bounds: IRange
  /// the rows of the block, from the rows the bounding copy kept (filtered rows are gone from it)
  rows(keptRows: readonly number[]): number[]
  columns(keptColumns: readonly number[]): number[]
}

export function multiSelectionLayout(ranges: readonly IRange[]): MultiSelectionLayout | null {
  if (ranges.length < 2) return null
  const first = ranges[0]!
  const sameRows = ranges.every(
    (range) => range.startRow === first.startRow && range.endRow === first.endRow,
  )
  const sameColumns = ranges.every(
    (range) => range.startColumn === first.startColumn && range.endColumn === first.endColumn,
  )
  if (!sameRows && !sameColumns) return null
  const bounds: IRange = {
    startRow: Math.min(...ranges.map((range) => range.startRow)),
    endRow: Math.max(...ranges.map((range) => range.endRow)),
    startColumn: Math.min(...ranges.map((range) => range.startColumn)),
    endColumn: Math.max(...ranges.map((range) => range.endColumn)),
  }
  const inAny = (index: number, axis: 'row' | 'column'): boolean =>
    ranges.some((range) =>
      axis === 'row'
        ? index >= range.startRow && index <= range.endRow
        : index >= range.startColumn && index <= range.endColumn,
    )
  return {
    bounds,
    // rows / columns of the block: those inside any area, in sheet order (an index covered by two
    // areas counts once)
    rows: (kept) => kept.filter((row) => inAny(row, 'row')),
    columns: (kept) => kept.filter((column) => inAny(column, 'column')),
  }
}
