import { describe, expect, it } from 'vitest'

import { applyUiTheme } from '../src/ui-theme'

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
