/**
 * Context-menu placement: keep the whole menu inside the window. A menu taller than the window
 * is capped to the window height (the .ctx-menu element then scrolls, like Office's long menus)
 * instead of overflowing off the bottom with no way to reach the last items.
 */
export const CTX_MENU_MARGIN = 4

export interface CtxMenuPlacement {
  x: number
  y: number
  /** Cap for the menu height; null when the whole menu fits */
  maxHeight: number | null
}

export function placeContextMenu(
  x: number,
  y: number,
  menuWidth: number,
  naturalHeight: number,
  viewportWidth: number,
  viewportHeight: number,
  margin = CTX_MENU_MARGIN,
): CtxMenuPlacement {
  const avail = Math.max(0, viewportHeight - margin * 2)
  const tooTall = naturalHeight > avail
  const height = tooTall ? avail : naturalHeight
  const px = Math.min(x, Math.max(0, viewportWidth - menuWidth - margin))
  // Prefer the click point; slide up just enough to fit; never above the top margin.
  const py = Math.max(margin, Math.min(y, viewportHeight - height - margin))
  return { x: Math.max(0, px), y: tooTall ? margin : py, maxHeight: tooTall ? avail : null }
}
