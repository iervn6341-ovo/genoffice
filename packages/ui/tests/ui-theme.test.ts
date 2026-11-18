import { describe, expect, it, vi } from 'vitest'

import { applyUiTheme, syncUiTheme, type UiThemeName } from '../src/ui-theme'

function fakeRoot() {
  const attrs = new Map<string, string>()
  return {
    attrs,
    setAttribute: (k: string, v: string) => void attrs.set(k, v),
    removeAttribute: (k: string) => void attrs.delete(k),
  } as unknown as HTMLElement & { attrs: Map<string, string> }
}

describe('applyUiTheme', () => {
  it('sets light and dark plainly', () => {
    const root = fakeRoot()
    applyUiTheme('light', root)
    expect(root.attrs.get('data-theme')).toBe('light')
    expect(root.attrs.has('data-theme-variant')).toBe(false)
    applyUiTheme('dark', root)
    expect(root.attrs.get('data-theme')).toBe('dark')
  })
  it('gray is the dark family plus a variant flag, so dark rules keep applying', () => {
    const root = fakeRoot()
    applyUiTheme('gray', root)
    expect(root.attrs.get('data-theme')).toBe('dark')
    expect(root.attrs.get('data-theme-variant')).toBe('gray')
  })
  it('leaving gray clears the variant; system clears both', () => {
    const root = fakeRoot()
    applyUiTheme('gray', root)
    applyUiTheme('dark', root)
    expect(root.attrs.has('data-theme-variant')).toBe(false)
    applyUiTheme('gray', root)
    applyUiTheme('system', root)
    expect(root.attrs.size).toBe(0)
  })
})

function themeSource() {
  let listener: ((theme: UiThemeName) => void) | undefined
  let resolve!: (theme: UiThemeName) => void
  let reject!: (error: Error) => void
  const pending = new Promise<UiThemeName>((yes, no) => {
    resolve = yes
    reject = no
  })
  return {
    getTheme: vi.fn(() => pending),
    onThemeChanged: vi.fn((handler: (theme: UiThemeName) => void) => {
      listener = handler
      return () => {
        listener = undefined
      }
    }),
    emit: (theme: UiThemeName) => listener?.(theme),
    resolve,
    reject,
  }
}

describe('theme synchronization during window bootstrap', () => {
  it('subscribes before querying the initial theme', async () => {
    const source = themeSource(),
      apply = vi.fn()
    const sync = syncUiTheme(source, apply)
    expect(source.onThemeChanged).toHaveBeenCalledOnce()
    await Promise.resolve()
    expect(source.getTheme).toHaveBeenCalledOnce()
    source.resolve('gray')
    await sync.ready
    expect(apply).toHaveBeenLastCalledWith('gray')
    sync.dispose()
  })
  it('keeps newer events instead of a stale startup response', async () => {
    const source = themeSource(),
      apply = vi.fn()
    const sync = syncUiTheme(source, apply)
    source.emit('light')
    source.emit('gray')
    source.resolve('dark')
    await sync.ready
    expect(apply.mock.calls).toEqual([['light'], ['gray']])
    source.emit('system')
    expect(apply).toHaveBeenLastCalledWith('system')
    sync.dispose()
  })
  it('updates while other bootstrap work is still pending', async () => {
    const source = themeSource(),
      root = fakeRoot()
    const sync = syncUiTheme(source, (value) => applyUiTheme(value, root))
    source.resolve('dark')
    await sync.ready
    // Language/font loading can finish later; no old snapshot is reapplied.
    source.emit('light')
    await Promise.resolve()
    expect(root.attrs.get('data-theme')).toBe('light')
    sync.dispose()
  })
  it('falls back to system only when no newer event exists', async () => {
    const source = themeSource(),
      apply = vi.fn()
    const sync = syncUiTheme(source, apply)
    source.emit('gray')
    source.reject(new Error('snapshot unavailable'))
    await sync.ready
    expect(apply.mock.calls).toEqual([['gray']])
    sync.dispose()
    const failed = themeSource()
    const fallback = syncUiTheme(failed, apply)
    failed.reject(new Error('standalone'))
    await fallback.ready
    expect(apply).toHaveBeenLastCalledWith('system')
    fallback.dispose()
  })
  it('stops updates when a settings panel closes, including pending snapshots', async () => {
    const source = themeSource(),
      apply = vi.fn()
    const sync = syncUiTheme(source, apply)
    sync.dispose()
    source.resolve('dark')
    source.emit('gray')
    await sync.ready
    expect(apply).not.toHaveBeenCalled()
  })
})
