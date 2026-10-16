import type { ProtectionRange, SelectionProtection } from './selection-protection'

export function canEditSelectionBorder(
  selection: Pick<SelectionProtection, 'range' | 'merges' | 'mergesComplete'> | null,
): boolean {
  return (
    selection !== null &&
    selection.mergesComplete !== false &&
    borderSelectionIsSafe(selection.range, selection.merges ?? [])
  )
}

export function mergeAt(
  merges: readonly ProtectionRange[],
  row: number,
  column: number,
): ProtectionRange | undefined {
  return merges.find(
    (merge) =>
      row >= merge.startRow &&
      row <= merge.endRow &&
      column >= merge.startColumn &&
      column <= merge.endColumn,
  )
}

export function borderSelectionIsSafe(
  range: ProtectionRange,
  merges: readonly ProtectionRange[],
): boolean {
  return merges.every(
    (merge) =>
      merge.startRow >= range.startRow &&
      merge.endRow <= range.endRow &&
      merge.startColumn >= range.startColumn &&
      merge.endColumn <= range.endColumn,
  )
}

export function visibleBorderSeam(
  merges: readonly ProtectionRange[],
  row: number,
  column: number,
  direction: 'down' | 'right',
): boolean {
  const here = mergeAt(merges, row, column)
  const there = mergeAt(
    merges,
    row + (direction === 'down' ? 1 : 0),
    column + (direction === 'right' ? 1 : 0),
  )
  return !here || here !== there
}

export function borderControlHasVisibleSeam(
  range: ProtectionRange,
  merges: readonly ProtectionRange[],
  direction: 'down' | 'right',
): boolean {
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (
        (direction === 'down' ? row < range.endRow : column < range.endColumn) &&
        visibleBorderSeam(merges, row, column, direction)
      )
        return true
    }
  }
  return false
}
