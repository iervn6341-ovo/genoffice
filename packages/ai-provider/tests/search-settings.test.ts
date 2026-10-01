import { describe, expect, it } from 'vitest'
import { defaultAiSettings, resolveAiSettings } from '../src/providers'
import {
  activeSearchProvider,
  defaultAiSearchSettings,
  resolveAiSearchSettings,
} from '../src/search-settings'

describe('search settings', () => {
  it('defaults to auto with empty keys and rides along in defaultAiSettings', () => {
    expect(defaultAiSearchSettings()).toEqual({
      provider: 'auto',
      providers: { serper: { apiKey: '' }, tavily: { apiKey: '' } },
    })
    expect(defaultAiSettings().search?.provider).toBe('auto')
    const resolved = resolveAiSettings(
      { provider: 'custom', providers: {} as never },
      defaultAiSettings(),
    )
    expect(resolved.search).toEqual(defaultAiSearchSettings())
  })

  it('merges and trims stored keys', () => {
    const s = resolveAiSearchSettings({
      provider: 'tavily',
      providers: { tavily: { apiKey: ' tvly-1 ' } } as never,
    })
    expect(s.provider).toBe('tavily')
    expect(s.providers.tavily.apiKey).toBe('tvly-1')
    expect(s.providers.serper.apiKey).toBe('')
  })

  it('activates a BYOK search provider only with a key', () => {
    expect(activeSearchProvider({ search: undefined })).toBe('auto')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: { serper: { apiKey: '' }, tavily: { apiKey: '' } },
        },
      }),
    ).toBe('auto')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: { serper: { apiKey: 'k' }, tavily: { apiKey: '' } },
        },
      }),
    ).toBe('serper')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: { serper: { apiKey: '   ' }, tavily: { apiKey: '' } },
        },
      }),
    ).toBe('auto')
    expect(activeSearchProvider({ search: { provider: 'bing', providers: {} } as never })).toBe(
      'auto',
    )
  })
})
