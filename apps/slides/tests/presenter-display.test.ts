/**
 * Presenter view: the audience window goes to the other screen, a real monitor or projector
 * rather than an iPad extending the desk through Sidecar, whatever order macOS lists them in.
 */
import { describe, it, expect, vi } from 'vitest'

vi.mock('electron', () => ({ BrowserWindow: {}, ipcMain: {}, screen: {} }))
vi.mock('@genoffice/electron-utils', () => ({ rendererUrl: () => '' }))
vi.mock('../src/main/session-state', () => ({
  runtime: {},
  sessions: new Map(),
  viewerWcIds: new Set(),
  windowRefs: {},
}))

const { pickAudienceDisplay } = await import('../src/main/presenter-show')

const display = (id: number, label: string, width: number, height: number) =>
  ({ id, label, bounds: { x: 0, y: 0, width, height } }) as Electron.Display

const builtIn = display(1, 'Built-in Retina Display', 1800, 1169)
const asus = display(2, 'ASUS VS229', 1920, 1080)
const sidecar = display(24, 'Sidecar Display (AirPlay)', 1334, 1000)

describe('pickAudienceDisplay', () => {
  it('picks the external monitor over Sidecar in either order', () => {
    expect(pickAudienceDisplay([builtIn, asus, sidecar], 1)?.id).toBe(2)
    expect(pickAudienceDisplay([builtIn, sidecar, asus], 1)?.id).toBe(2)
  })

  it('never picks the presenter screen', () => {
    expect(pickAudienceDisplay([builtIn, asus, sidecar], 2)?.id).toBe(1)
    expect(pickAudienceDisplay([builtIn], 1)).toBeUndefined()
  })

  it('uses Sidecar when it is the only other screen', () => {
    expect(pickAudienceDisplay([builtIn, sidecar], 1)?.id).toBe(24)
  })

  it('honours the Monitor picked in Slide Show → Monitors, unless it is the presenter screen', () => {
    expect(pickAudienceDisplay([builtIn, asus, sidecar], 1, 'Sidecar Display (AirPlay)')?.id).toBe(
      24,
    )
    expect(pickAudienceDisplay([builtIn, asus, sidecar], 1, 'Built-in Retina Display')?.id).toBe(2)
    expect(pickAudienceDisplay([builtIn, asus, sidecar], 1, 'Unplugged Projector')?.id).toBe(2)
  })

  it('prefers the larger of two monitors', () => {
    const projector = display(3, 'Projector', 1280, 720)
    expect(pickAudienceDisplay([builtIn, projector, asus], 1)?.id).toBe(2)
  })
})
