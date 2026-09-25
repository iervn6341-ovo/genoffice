import type { MouseEvent as ReactMouseEvent } from 'react'

/**
 * Ribbon buttons must not take focus from the document: a mousedown that moves focus collapses
 * the editor selection, so the command then has nothing to act on. Attach as
 * `onMouseDownCapture` on the ribbon root (capture, because panels stop propagation of their
 * own mousedown). Text fields, selects and labels (native color pickers) keep normal focus
 * behavior; clicks still fire, only the focus change is suppressed.
 */
export function keepEditorFocusOnRibbonPress(e: ReactMouseEvent<HTMLElement>): void {
  const target = e.target
  if (!(target instanceof Element) || !target.closest('button')) return
  if (target.closest('input, textarea, select, label, [contenteditable="true"]')) return
  e.preventDefault()
}
