/**
 * After a selected object (chart, shape, picture) is deleted its DOM node, which held keyboard
 * focus, disappears and focus falls to <body>: the next ⌘Z / arrow key reaches nothing until the
 * grid is clicked. Excel keeps the keyboard in the workbook, so hand focus back to the grid - but
 * only when focus really was lost (never take it from the AI composer or a dialog field).
 */
let handler: (() => void) | null = null

/** The app registers how to focus the grid (Univer's hidden editor host). */
export function setGridFocusHandler(next: (() => void) | null): void {
  handler = next
}

/// The removed node can still hold focus for a tick (the React commit that unmounts it lands
/// after the selection is cleared), so look again a few times; the first check that finds focus
/// lost wins and ends the watch.
export const GRID_FOCUS_CHECKS_MS: readonly number[] = [0, 120, 400]

export function scheduleGridFocusReturn(
  doc: Document = document,
  checksMs: readonly number[] = GRID_FOCUS_CHECKS_MS,
): void {
  let done = false
  for (const delay of checksMs) {
    setTimeout(() => {
      if (done) return
      const active = doc.activeElement
      if (active && active !== doc.body) return
      done = true
      handler?.()
    }, delay)
  }
}
