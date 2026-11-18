import { beforeEach, describe, expect, it, vi } from 'vitest'

const mock = vi.hoisted(() => ({
  handles: new Map<string, (...args: any[]) => any>(),
  sessions: new Map<number, unknown>(),
  displays: [{ id: 1, label: 'Presenter', bounds: { x: 0, y: 0, width: 1200, height: 800 } }],
  windows: [] as any[],
}))
vi.mock('electron', () => ({
  BrowserWindow: class {
    static fromWebContents() {
      return null
    }
    webContents = { id: 200, isDestroyed: () => false, send: vi.fn(), focus: vi.fn() }
    once = vi.fn()
    on = vi.fn()
    loadURL = vi.fn(async () => {})
    isDestroyed = () => false
    isSimpleFullScreen = () => false
    close = vi.fn()
    constructor(options: unknown) {
      mock.windows.push({ options, win: this })
    }
  },
  ipcMain: {
    handle: (name: string, fn: (...args: any[]) => any) => mock.handles.set(name, fn),
    on: vi.fn(),
  },
  screen: { getAllDisplays: () => mock.displays, getPrimaryDisplay: () => mock.displays[0] },
}))
vi.mock('@genoffice/electron-utils', () => ({
  rendererUrl: () => 'genoffice-app://slides/?mode=audience',
}))
vi.mock('../src/main/session-state', () => ({
  runtime: { preloadPath: 'preload.js' },
  sessions: mock.sessions,
  viewerWcIds: new Set(),
  windowRefs: { shellWindow: null },
}))
import { registerPresenterIpc } from '../src/main/presenter-show'
beforeEach(() => {
  mock.windows.length = 0
  mock.sessions.clear()
  mock.displays.length = 1
  registerPresenterIpc()
})

describe('multi-display playback scheduling', () => {
  it.each([true, false])(
    'keeps the presenter rendering and restores the previous throttling policy (%s)',
    (policy) => {
      const contents = {
        id: 100,
        getBackgroundThrottling: () => policy,
        setBackgroundThrottling: vi.fn(),
        isDestroyed: () => false,
        once: vi.fn(),
      }
      const event = { sender: contents }
      expect(mock.handles.get('slides:presenter-start')!(event)).toEqual({ audience: false })
      expect(contents.setBackgroundThrottling).toHaveBeenCalledWith(false)
      mock.handles.get('slides:presenter-end')!(event)
      expect(contents.setBackgroundThrottling).toHaveBeenLastCalledWith(policy)
    },
  )
  it('disables throttling on the audience window too and restores the presenter when ending', () => {
    mock.displays.push({
      id: 2,
      label: 'Projector',
      bounds: { x: 1200, y: 0, width: 1920, height: 1080 },
    })
    const contents = {
      id: 101,
      getBackgroundThrottling: () => true,
      setBackgroundThrottling: vi.fn(),
      isDestroyed: () => false,
      once: vi.fn(),
    }
    mock.sessions.set(contents.id, {})
    const event = { sender: contents }
    expect(mock.handles.get('slides:presenter-start')!(event)).toEqual({ audience: true })
    expect(mock.windows[0].options.webPreferences.backgroundThrottling).toBe(false)
    // Re-entering must not forget the editor's original true policy.
    mock.handles.get('slides:presenter-start')!(event)
    mock.handles.get('slides:presenter-end')!(event)
    expect(contents.setBackgroundThrottling).toHaveBeenLastCalledWith(true)
    expect(mock.windows[0].win.close).toHaveBeenCalledOnce()
  })
})
