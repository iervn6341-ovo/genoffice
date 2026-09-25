/**
 * Presenter-view multi-screen show, extracted from slides-main.ts: audience
 * window on the external display + presenter->audience state forwarding.
 * The audience window shares the same Session as the presenter (sessions holds
 * one reference per wc.id), so read-only IPC (get-render-slides/get-animations/
 * get-transition…) works naturally, and it sees the in-memory document
 * (including unsaved changes) without re-reading from disk.
 */
import { rendererUrl } from '@genoffice/electron-utils'
import { BrowserWindow, ipcMain, screen } from 'electron'
import type { WebContents } from 'electron'
import type { AudienceNavAction, ShowInkEvent, ShowSyncState } from '../shared/ipc'
import { runtime, sessions, viewerWcIds, windowRefs } from './session-state'

interface PresenterShow {
  presenterWc: WebContents
  audienceWin: BrowserWindow | null
  /** Most recent sync state: audience-ready re-sends it when the audience window mounts after the broadcast */
  lastSync: ShowSyncState | null
}
const presenterShows = new Map<number, PresenterShow>()
/** Audience wc.id -> presenter wc.id (routing for navigation actions sent back) */
const audiencePresenter = new Map<number, number>()

/** Host window of the presenter renderer (in tab mode, the shell's single window) */
function presenterHostWindow(wc: WebContents): BrowserWindow | null {
  return BrowserWindow.fromWebContents(wc) ?? windowRefs.shellWindow
}

/** Fullscreen a window on a given display. macOS uses simpleFullScreen to avoid Spaces animations and cross-screen move restrictions */
function fullScreenOnDisplay(win: BrowserWindow, bounds: Electron.Rectangle): void {
  if (process.platform === 'darwin') {
    if (win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
    win.setBounds(bounds)
    win.setSimpleFullScreen(true)
  } else {
    if (win.isFullScreen()) win.setFullScreen(false)
    win.setBounds(bounds)
    win.setFullScreen(true)
  }
}

/**
 * The screen the audience sees (PowerPoint's "Automatic" monitor): any display other than the
 * presenter's, preferring a real monitor/projector over an iPad used as a Sidecar extension of
 * the desk (display order is not stable across reconnects), then the larger screen.
 */
export function pickAudienceDisplay(
  displays: Electron.Display[],
  hostId: number,
  /** Slide Show → Monitor: a display label the user picked; ignored when absent or on the presenter's screen */
  preferred?: string | null,
): Electron.Display | undefined {
  const chosen = preferred ? displays.find((d) => d.label === preferred && d.id !== hostId) : null
  if (chosen) return chosen
  const sidecar = (d: Electron.Display) => /\bsidecar\b/i.test(d.label ?? '')
  return displays
    .filter((d) => d.id !== hostId)
    .sort(
      (a, b) =>
        Number(sidecar(a)) - Number(sidecar(b)) ||
        b.bounds.width * b.bounds.height - a.bounds.width * a.bounds.height,
    )[0]
}

function closePresenterShow(presenterId: number): void {
  const show = presenterShows.get(presenterId)
  if (!show) return
  presenterShows.delete(presenterId)
  const win = show.audienceWin
  show.audienceWin = null
  if (win && !win.isDestroyed()) {
    if (process.platform === 'darwin' && win.isSimpleFullScreen()) win.setSimpleFullScreen(false)
    win.close()
  }
}

/** Register the slides:presenter-* / slides:audience-* channels (called from registerSlidesIpc). */
export function registerPresenterIpc(): void {
  // Show start: with a second screen and "Use Presenter View" on, the show opens presenter view
  ipcMain.handle('slides:display-count', () => screen.getAllDisplays().length)
  ipcMain.handle('slides:display-list', () => {
    const primary = screen.getPrimaryDisplay().id
    return screen
      .getAllDisplays()
      .map((d) => ({ id: d.id, label: d.label || `Display ${d.id}`, primary: d.id === primary }))
  })

  ipcMain.handle('slides:presenter-start', (e, opts?: { monitor?: string | null }) => {
    const existing = presenterShows.get(e.sender.id)
    if (existing) return { audience: existing.audienceWin != null }
    const show: PresenterShow = { presenterWc: e.sender, audienceWin: null, lastSync: null }
    presenterShows.set(e.sender.id, show)
    e.sender.once('destroyed', () => closePresenterShow(e.sender.id))

    const session = sessions.get(e.sender.id)
    const host = presenterHostWindow(e.sender)
    const hostDisplay = host
      ? screen.getDisplayMatching(host.getBounds())
      : screen.getPrimaryDisplay()
    const external = pickAudienceDisplay(screen.getAllDisplays(), hostDisplay.id, opts?.monitor)
    if (!external || !session) return { audience: false }

    const win = new BrowserWindow({
      ...external.bounds,
      frame: false,
      show: false,
      backgroundColor: '#000000',
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    show.audienceWin = win
    sessions.set(win.webContents.id, session)
    audiencePresenter.set(win.webContents.id, e.sender.id)
    const audienceWcId = win.webContents.id
    viewerWcIds.add(audienceWcId)
    win.once('ready-to-show', () => {
      // PowerPoint keeps the keyboard on the presenter: show the audience screen without
      // activating it, then hand focus back (a plain show() made the audience window key
      // while no page in it held focus, so Esc and the arrow keys went nowhere)
      win.showInactive()
      fullScreenOnDisplay(win, external.bounds)
      if (host && !host.isDestroyed()) host.focus()
      if (!e.sender.isDestroyed()) e.sender.focus()
    })
    // A click on the audience screen makes it the key window: give its page the keyboard too
    win.on('focus', () => {
      if (!win.webContents.isDestroyed()) win.webContents.focus()
    })
    win.on('closed', () => {
      sessions.delete(audienceWcId)
      viewerWcIds.delete(audienceWcId)
      audiencePresenter.delete(audienceWcId)
      const s = presenterShows.get(e.sender.id)
      if (s?.audienceWin === win) s.audienceWin = null
    })
    void win.loadURL(rendererUrl(runtime.rendererDevUrl, 'slides', { mode: 'audience' }))
    return { audience: true }
  })

  ipcMain.on('slides:presenter-sync', (e, state: ShowSyncState) => {
    const show = presenterShows.get(e.sender.id)
    if (!show) return
    show.lastSync = state
    const wc = show.audienceWin?.webContents
    if (wc && !wc.isDestroyed()) wc.send('slides:show-sync', state)
  })

  ipcMain.on('slides:presenter-ink', (e, ev: ShowInkEvent) => {
    const wc = presenterShows.get(e.sender.id)?.audienceWin?.webContents
    if (wc && !wc.isDestroyed()) wc.send('slides:show-ink', ev)
  })

  // Ink drawn with the mouse on the audience screen (laser, pen…) mirrors back to the presenter
  ipcMain.on('slides:audience-ink', (e, ev: ShowInkEvent) => {
    const pid = audiencePresenter.get(e.sender.id)
    const wc = pid != null ? presenterShows.get(pid)?.presenterWc : null
    if (wc && !wc.isDestroyed()) wc.send('slides:show-ink', ev)
  })

  ipcMain.handle('slides:presenter-swap', (e) => {
    const show = presenterShows.get(e.sender.id)
    const aWin = show?.audienceWin
    const host = presenterHostWindow(e.sender)
    if (!show || !aWin || aWin.isDestroyed() || !host) return false
    const aDisplay = screen.getDisplayMatching(aWin.getBounds())
    const hDisplay = screen.getDisplayMatching(host.getBounds())
    if (aDisplay.id === hDisplay.id) return false
    fullScreenOnDisplay(aWin, hDisplay.bounds)
    // Move the presenter window to the audience's former screen, keeping its fullscreen form (HTML fullscreen uses native fullscreen)
    if (host.isFullScreen()) {
      host.once('leave-full-screen', () => {
        host.setBounds(aDisplay.workArea)
        host.setFullScreen(true)
      })
      host.setFullScreen(false)
    } else if (host.isSimpleFullScreen()) {
      host.setSimpleFullScreen(false)
      host.setBounds(aDisplay.workArea)
      host.setSimpleFullScreen(true)
    } else {
      host.setBounds(aDisplay.workArea)
    }
    return true
  })

  ipcMain.handle('slides:presenter-end', (e) => {
    closePresenterShow(e.sender.id)
  })

  ipcMain.handle('slides:audience-ready', (e) => {
    const pid = audiencePresenter.get(e.sender.id)
    return (pid != null ? presenterShows.get(pid)?.lastSync : null) ?? null
  })

  ipcMain.on('slides:audience-nav', (e, action: AudienceNavAction) => {
    const pid = audiencePresenter.get(e.sender.id)
    const wc = pid != null ? presenterShows.get(pid)?.presenterWc : null
    if (wc && !wc.isDestroyed()) wc.send('slides:audience-nav', action)
  })
}
