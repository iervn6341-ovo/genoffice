import { describe, expect, it } from 'vitest'

import { DEFAULT_AI_PANEL_PREFS, normalizeAiPanelPrefs } from '../src/ai-panel-prefs'
import { resolveContentDark } from '../src/content-theme'

describe('resolveContentDark', () => {
  it('auto follows the app theme; a pinned choice ignores it', () => {
    expect(resolveContentDark('auto', true)).toBe(true)
    expect(resolveContentDark('auto', false)).toBe(false)
    expect(resolveContentDark('dark', false)).toBe(true)
    expect(resolveContentDark('light', true)).toBe(false)
  })
})

describe('contentTheme preference', () => {
  it('defaults to auto and rejects unknown values', () => {
    expect(DEFAULT_AI_PANEL_PREFS.contentTheme).toBe('auto')
    expect(normalizeAiPanelPrefs({ contentTheme: 'sepia' }).contentTheme).toBe('auto')
    expect(normalizeAiPanelPrefs({ contentTheme: 'dark' }).contentTheme).toBe('dark')
    expect(normalizeAiPanelPrefs({ contentTheme: 'light' }).contentTheme).toBe('light')
    expect(normalizeAiPanelPrefs(undefined).contentTheme).toBe('auto')
  })
})
