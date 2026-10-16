// @vitest-environment jsdom
import { createElement, useEffect } from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { applyAiPanelPrefs } from '@genoffice/ui'
import { uiThemeIsDark, useContentDark } from '../src/renderer/content-theme'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function mockSystemDark(matches: boolean): void {
  window.matchMedia = vi.fn().mockReturnValue({
    matches,
    addEventListener: () => undefined,
    removeEventListener: () => undefined,
  }) as unknown as typeof window.matchMedia
}

afterEach(() => {
  document.documentElement.removeAttribute('data-theme')
  delete (window as { desktopApi?: unknown }).desktopApi
})

describe('uiThemeIsDark', () => {
  it('reads explicit data-theme, else the OS appearance', () => {
    mockSystemDark(true)
    document.documentElement.setAttribute('data-theme', 'light')
    expect(uiThemeIsDark()).toBe(false)
    document.documentElement.setAttribute('data-theme', 'dark')
    expect(uiThemeIsDark()).toBe(true)
    document.documentElement.removeAttribute('data-theme')
    expect(uiThemeIsDark()).toBe(true)
    mockSystemDark(false)
    expect(uiThemeIsDark()).toBe(false)
  })
})

describe('useContentDark', () => {
  it('follows the app theme until pinned, then keeps the saved choice across theme changes', () => {
    mockSystemDark(false)
    document.documentElement.setAttribute('data-theme', 'light')
    let notify: () => void = () => undefined
    const saved: unknown[] = []
    // the shell's shared preferences record: every save is echoed back to the renderer store
    ;(window as unknown as { desktopApi: unknown }).desktopApi = {
      onThemeChanged: (cb: () => void) => {
        notify = cb
        return () => undefined
      },
      setAiPanelPrefs: (patch: Record<string, unknown>) => {
        saved.push(patch)
        applyAiPanelPrefs({ ...patch })
      },
    }
    applyAiPanelPrefs({ contentTheme: 'auto' })
    let state: readonly [boolean, (d: boolean) => void] = [false, () => undefined]
    function Probe(): null {
      const s = useContentDark()
      useEffect(() => {
        state = s
      })
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(Probe)))
    expect(state[0]).toBe(false)

    // auto: a theme change carries the content with it
    document.documentElement.setAttribute('data-theme', 'dark')
    act(() => notify())
    expect(state[0]).toBe(true)

    // pinning light on a dark app theme is saved and survives the next theme change
    act(() => state[1](false))
    expect(saved.at(-1)).toEqual({ contentTheme: 'light' })
    expect(state[0]).toBe(false)
    document.documentElement.setAttribute('data-theme', 'light')
    act(() => notify())
    document.documentElement.setAttribute('data-theme', 'dark')
    act(() => notify())
    expect(state[0]).toBe(false)

    act(() => state[1](true))
    expect(saved.at(-1)).toEqual({ contentTheme: 'dark' })
    expect(state[0]).toBe(true)
    expect(document.documentElement.getAttribute('data-theme')).toBe('dark')

    // Settings can put it back on "follow theme"
    act(() => applyAiPanelPrefs({ contentTheme: 'auto' }))
    document.documentElement.setAttribute('data-theme', 'light')
    act(() => notify())
    expect(state[0]).toBe(false)
    act(() => root.unmount())
  })
})
