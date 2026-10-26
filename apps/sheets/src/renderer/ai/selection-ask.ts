export interface Point {
  x: number
  y: number
}

export interface ViewportBounds {
  left: number
  top: number
  right: number
  bottom: number
}

export interface SelectionAskPosition {
  left: number
  top: number
}

export interface SelectionAskAnchor {
  pointer: Point
  bounds: ViewportBounds
}

export const SELECTION_DRAG_THRESHOLD_PX = 4

/** Ignore pointer jitter so an ordinary cell click never summons Ask AI. */
export function isSelectionDrag(
  start: Point,
  current: Point,
  threshold = SELECTION_DRAG_THRESHOLD_PX,
): boolean {
  return Math.hypot(current.x - start.x, current.y - start.y) >= threshold
}

/** How far below the drag endpoint a trigger must start to clear the fill handle (Excel has none). */
export const FILL_HANDLE_CLEARANCE_PX = 30

/**
 * Keep the trigger beside the drag endpoint and inside the grid viewport. The endpoint sits in the
 * selection's last cell, so the fill handle (its bottom-right corner) lies below and to the right
 * of it: the trigger goes ABOVE-right of the pointer, and only when there is no room up there does
 * it drop below - then past the handle's row - so it never covers the handle.
 */
export function selectionAskPosition(
  pointer: Point,
  bounds: ViewportBounds,
  width: number,
  height: number,
  gap = 8,
): SelectionAskPosition {
  const inset = 4
  const minLeft = bounds.left + inset
  const maxLeft = Math.max(minLeft, bounds.right - width - inset)
  const minTop = bounds.top + inset
  const maxTop = Math.max(minTop, bounds.bottom - height - inset)
  const preferredLeft = pointer.x + gap
  const flippedLeft = pointer.x - width - gap
  const above = pointer.y - height - gap
  const below = pointer.y + FILL_HANDLE_CLEARANCE_PX
  const top = above >= minTop ? above : below + height <= bounds.bottom - inset ? below : above

  return {
    left: Math.min(
      Math.max(
        preferredLeft + width <= bounds.right - inset ? preferredLeft : flippedLeft,
        minLeft,
      ),
      maxLeft,
    ),
    top: Math.min(Math.max(top, minTop), maxTop),
  }
}
