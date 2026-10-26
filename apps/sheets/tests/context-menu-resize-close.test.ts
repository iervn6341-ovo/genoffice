// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'

import { installContextMenuResizeClose } from '../src/renderer/context-menu-resize-close'

describe('installContextMenuResizeClose', () => {
  it('hides a visible menu on window resize', () => {
    const menu = { visible: true, hideContextMenu: vi.fn() }
    const h = installContextMenuResizeClose(menu)
    window.dispatchEvent(new Event('resize'))
    expect(menu.hideContextMenu).toHaveBeenCalledTimes(1)
    h.dispose()
  })
  it('does nothing when no menu is open', () => {
    const menu = { visible: false, hideContextMenu: vi.fn() }
    const h = installContextMenuResizeClose(menu)
    window.dispatchEvent(new Event('resize'))
    expect(menu.hideContextMenu).not.toHaveBeenCalled()
    h.dispose()
  })
  it('stops after dispose', () => {
    const menu = { visible: true, hideContextMenu: vi.fn() }
    installContextMenuResizeClose(menu).dispose()
    window.dispatchEvent(new Event('resize'))
    expect(menu.hideContextMenu).not.toHaveBeenCalled()
  })
})
