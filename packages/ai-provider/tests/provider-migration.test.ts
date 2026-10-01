import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  activeProvider,
  defaultAiSettings,
  resolveAiSettings,
} from '../src/providers'

describe('provider selection without an account fallback', () => {
  it('keeps an incomplete explicit provider selected instead of sending data to another service', () => {
    const stored = defaultAiSettings({ anthropic: 'configured-key' })
    stored.provider = 'custom'
    const result = resolveAiSettings(stored, defaultAiSettings())
    expect(activeProvider(result)).toBe('custom')
    expect(result.providers.anthropic.apiKey).toBe('configured-key')
  })

  it('migrates a removed provider to an existing configured service and discards its credentials', () => {
    const result = resolveAiSettings(
      {
        provider: 'retired-service',
        providers: {
          'retired-service': { apiKey: 'old-private-key', model: 'old-model' },
          openai: {
            apiKey: ' existing-key ',
            model: 'saved-model',
            baseUrl: 'https://gateway.example/v1',
          },
        },
      } as never,
      defaultAiSettings(),
    )
    expect(result.provider).toBe('openai')
    expect(result.providers.openai).toMatchObject({
      apiKey: 'existing-key',
      model: 'saved-model',
      baseUrl: 'https://gateway.example/v1',
    })
    expect(JSON.stringify(result)).not.toContain('old-private-key')
    expect(Object.keys(result.providers).sort()).toEqual(AI_PROVIDERS.map((p) => p.id).sort())
  })

  it('preserves Codex auto-discovery when explicitly selected', () => {
    const result = resolveAiSettings({ provider: 'codex' }, defaultAiSettings())
    expect(activeProvider(result)).toBe('codex')
  })

  it('recognizes a configured keyless local endpoint when migrating', () => {
    const result = resolveAiSettings(
      {
        provider: 'retired-service',
        providers: {
          custom: { baseUrl: 'http://localhost:11434/v1', model: 'local-model', apiKey: '' },
        },
      } as never,
      defaultAiSettings(),
    )
    expect(result.provider).toBe('custom')
    expect(result.providers.custom.apiKey).toBe('')
    expect(result.providers.custom.model).toBe('local-model')
  })

  it('leaves a fresh profile unconfigured and never implicitly activates CLI authentication', () => {
    const result = resolveAiSettings({}, defaultAiSettings())
    expect(result.provider).toBe('custom')
    expect(result.providers.custom.baseUrl).toBe('')
  })

  it('migrates legacy keyless endpoint settings and preserves the model', () => {
    const result = resolveAiSettings(
      { baseUrl: 'http://localhost:1234/v1', model: ' local ' },
      defaultAiSettings(),
    )
    expect(result.provider).toBe('custom')
    expect(result.providers.custom.model).toBe('local')
  })
})
