/**
 * Resizing the window while a context menu is open closes the menu (Excel).
 *
 * Univer re-runs its clamp on every window resize: a menu opened near the
 * bottom edge is cut to the space below it, and enlarging the window then
 * stretches it to its full height (and shrinking it flips/shifts the menu
 * and its submenus). Closing the menu on resize avoids the menu changing
 * size under the pointer.
 */
export interface ContextMenuHost {
  readonly visible: boolean
  hideContextMenu(): void
}

export function installContextMenuResizeClose(
  menu: ContextMenuHost,
  target: Pick<Window, 'addEventListener' | 'removeEventListener'> = window,
): { dispose(): void } {
  const onResize = () => {
    if (menu.visible) menu.hideContextMenu()
  }
  // capture + first in line: runs before Univer's own resize listeners re-clamp
  target.addEventListener('resize', onResize, true)
  return { dispose: () => target.removeEventListener('resize', onResize, true) }
}
