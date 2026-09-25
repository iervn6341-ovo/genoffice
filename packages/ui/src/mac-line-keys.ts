const IS_MAC = typeof navigator !== 'undefined' && navigator.platform.toLowerCase().includes('mac')

/**
 * Home / End in rich-text editors on macOS. Chromium maps both keys to "scroll to the
 * beginning / end of the document" there and leaves the caret where it was, while Word and
 * PowerPoint for Mac move it to the start / end of the visual line (⇧ extends the selection,
 * ⌘ goes to the start / end of the document) — Windows already behaves that way natively.
 * Returns true when the key was handled. Editors that bind these keys themselves (CodeMirror,
 * the spreadsheet grid) have already called preventDefault and are left alone.
 */
export function handleMacLineBoundaryKey(e: KeyboardEvent): boolean {
  if (!IS_MAC || e.defaultPrevented || e.altKey || e.ctrlKey) return false
  if (e.key !== 'Home' && e.key !== 'End') return false
  const target = e.target
  if (!(target instanceof HTMLElement) || !target.isContentEditable) return false
  const sel = window.getSelection()
  if (!sel || sel.rangeCount === 0) return false
  sel.modify(
    e.shiftKey ? 'extend' : 'move',
    e.key === 'Home' ? 'backward' : 'forward',
    e.metaKey ? 'documentboundary' : 'lineboundary',
  )
  e.preventDefault()
  return true
}

/** Install {@link handleMacLineBoundaryKey} for the whole window (bubble phase, after editors) */
export function installMacLineBoundaryKeys(): void {
  if (!IS_MAC || typeof window === 'undefined') return
  window.addEventListener('keydown', (e) => {
    handleMacLineBoundaryKey(e)
  })
}
