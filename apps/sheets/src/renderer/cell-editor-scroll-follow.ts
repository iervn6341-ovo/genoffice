/**
 * Scrolling the grid while a cell is being edited must carry the in-cell editor with its cell, like
 * Excel (the editor scrolls away with the cell, it does not stay behind over other cells).
 *
 * Univer places the editor overlay once, when editing starts, and again only on zoom/resize;
 * a scroll never reaches it, so the overlay stays at its first viewport position while the grid
 * moves underneath (a tall merged cell's editor ends up hovering over unrelated rows).
 * `refreshEditCellPosition(false)` recomputes the editor layout from the current scroll offset and
 * `fitTextSize()` pushes it to the overlay — the same two calls Univer makes when the editor opens —
 * so they are replayed on every scroll while the editor is visible, coalesced to one per frame.
 */
export interface EditorBridgeLike {
  isVisible(): { visible: boolean }
  refreshEditCellPosition(resetSizeOnly: boolean): void
}

export interface EditorResizeLike {
  fitTextSize(): void
}

export interface ScrollFollowHost {
  bridge: EditorBridgeLike
  resize: EditorResizeLike
  /** Subscribe to grid scrolls; returns the unsubscribe */
  onScroll(handler: () => void): () => void
  schedule?: (cb: () => void) => void
}

export function installCellEditorScrollFollow(host: ScrollFollowHost): () => void {
  const schedule =
    host.schedule ??
    ((cb: () => void) => {
      if (typeof requestAnimationFrame === 'function') requestAnimationFrame(cb)
      else setTimeout(cb, 0)
    })
  let pending = false
  const unsubscribe = host.onScroll(() => {
    if (pending || !host.bridge.isVisible().visible) return
    pending = true
    schedule(() => {
      pending = false
      // the editor may have closed between the scroll and the frame
      if (!host.bridge.isVisible().visible) return
      try {
        host.bridge.refreshEditCellPosition(false)
        host.resize.fitTextSize()
      } catch {
        // the edit cell was torn down mid-frame
      }
    })
  })
  return unsubscribe
}
