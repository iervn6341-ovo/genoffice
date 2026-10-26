import { isGridKeyTarget, type ClearSelectionKeyEvent } from './clear-selection-keyboard'

/**
 * Ctrl/Cmd+Z and Ctrl/Cmd+Shift+Z / Ctrl+Y must undo and redo after a ribbon
 * action too (Excel keeps keyboard focus in the workbook). Univer binds both
 * shortcuts only while its own grid host has focus; after a click on a ribbon
 * button (Format as Table, Bold, ...) focus stays on that button or the page
 * body, the keypress reached nothing, and the table could not be undone from
 * the keyboard. Keys aimed at Univer's container are left to Univer (no double
 * undo); text fields and dialogs keep their own native undo.
 */
export type UndoShortcutEvent = Pick<
  ClearSelectionKeyEvent,
  'key' | 'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'defaultPrevented' | 'isComposing' | 'target'
>

const SHEET_CONTAINER_SELECTOR = '#univer-container'

export function undoShortcutAction(
  event: UndoShortcutEvent,
  isCellEditing: boolean,
): 'undo' | 'redo' | null {
  if (!(event.metaKey || event.ctrlKey) || event.altKey) return null
  if (event.defaultPrevented || event.isComposing || isCellEditing) return null
  const key = event.key.toLowerCase()
  let action: 'undo' | 'redo' | null = null
  if (key === 'z') action = event.shiftKey ? 'redo' : 'undo'
  else if (key === 'y' && !event.shiftKey) action = 'redo'
  if (!action) return null
  const target = event.target as { closest?(selector: string): unknown } | null
  if (target && typeof target.closest === 'function') {
    if (target.closest(SHEET_CONTAINER_SELECTOR)) return null
  }
  return isGridKeyTarget(event.target) ? action : null
}
