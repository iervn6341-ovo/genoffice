import type { BrowserWindow } from 'electron'

/** Dismiss chrome once per native move/resize burst, without touching the frame. */
export function installNativeFrameEvents(win: BrowserWindow, dismiss: () => void): void {
  let quietTimer: ReturnType<typeof setTimeout> | null = null
  const onGesture = (): void => {
    const starting = quietTimer === null
    if (quietTimer) clearTimeout(quietTimer)
    quietTimer = setTimeout(() => {
      quietTimer = null
    }, 150)
    if (starting) dismiss()
  }
  win.on('will-move', onGesture)
  win.on('will-resize', onGesture)
  win.once('closed', () => {
    if (quietTimer) clearTimeout(quietTimer)
    quietTimer = null
    win.off('will-move', onGesture)
    win.off('will-resize', onGesture)
  })
}
