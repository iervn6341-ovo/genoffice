/**
 * Keep context menus and their submenus inside the window.
 *
 * Main menu: opens down-right from the pointer; when it would run off the bottom it flips up
 * (Word/macOS), and when neither fits it is pinned to the window and scrolls (`maxHeight`).
 * Submenu: opens to the right of its parent item; when there is no room it flips to the left, and
 * is shifted up (or capped and scrolled) so it never leaves the window vertically.
 */
export const MENU_MARGIN = 8

export interface MenuPlacement {
  left: number
  top: number
  /** set only when the menu is taller than the window (it then scrolls) */
  maxHeight?: number
}

export function placeContextMenu(
  pointer: { x: number; y: number },
  size: { width: number; height: number },
  viewport: { width: number; height: number },
  margin = MENU_MARGIN,
): MenuPlacement {
  const room = viewport.height - margin * 2
  const maxHeight = size.height > room ? Math.max(room, 0) : undefined
  const height = maxHeight ?? size.height
  let left = pointer.x
  if (left + size.width > viewport.width - margin) left = viewport.width - size.width - margin
  left = Math.max(margin, left)
  let top = pointer.y
  if (top + height > viewport.height - margin) {
    const flipped = pointer.y - height
    top = flipped >= margin ? flipped : viewport.height - margin - height
  }
  top = Math.max(margin, top)
  return maxHeight === undefined ? { left, top } : { left, top, maxHeight }
}

export function placeSubmenu(
  anchor: { left: number; right: number; top: number },
  size: { width: number; height: number },
  viewport: { width: number; height: number },
  margin = MENU_MARGIN,
  overlap = 4,
  lift = 5,
): MenuPlacement {
  let left = anchor.right - overlap
  if (left + size.width > viewport.width - margin) {
    left = anchor.left - size.width + overlap
    if (left < margin) left = Math.max(margin, viewport.width - margin - size.width)
  }
  const room = viewport.height - margin * 2
  const maxHeight = size.height > room ? Math.max(room, 0) : undefined
  const height = maxHeight ?? size.height
  let top = anchor.top - lift
  if (top + height > viewport.height - margin) top = viewport.height - margin - height
  top = Math.max(margin, top)
  return maxHeight === undefined ? { left, top } : { left, top, maxHeight }
}
