import type { Selection } from '@tiptap/pm/state'

/**
 * Is document position `pos` inside the selection? Checks every range: a table CellSelection has
 * one range per cell (its `from`/`to` cover only the head cell), so a right-click inside the
 * dragged cell range must count as "inside" or the menu would collapse the range to a caret.
 */
export function posInsideSelection(selection: Pick<Selection, 'ranges'>, pos: number): boolean {
  return selection.ranges.some((r) => pos >= r.$from.pos && pos <= r.$to.pos)
}
