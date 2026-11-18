import { createRequire } from 'node:module'
import { join } from 'node:path'
import type { BrowserWindow } from 'electron'

/** Install on every shell: both detach and merge can stall the next AppKit drag. */
export function installWindowGestureGuard(win: BrowserWindow): void {
  if (process.platform !== 'darwin' || win.isDestroyed()) return
  const version = process.getSystemVersion().split('.').map(Number)
  if (version[0] !== 27 || version[1] !== 2) return
  try {
    const require = createRequire(join(__dirname, 'index.js'))
    const native = require(join(__dirname, '../native/window-gesture-guard.node')) as {
      attach(handle: Buffer): boolean
    }
    native.attach(win.getNativeWindowHandle())
  } catch (error) {
    console.warn('[window] Native drag recovery unavailable:', error)
  }
}
