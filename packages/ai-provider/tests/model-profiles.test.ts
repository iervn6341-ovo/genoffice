import { describe, expect, it } from 'vitest'
import { defaultAiSettings, resolveAiSettings } from '../src/providers'
import { activateModelProfile, removeModelProfile, saveModelProfile } from '../src/model-profiles'
import { registerAiSettingsHandlers } from '../src/settings-ipc'
import type { AiSettings } from '../src/types'

describe('saved model profiles', () => {
  it('migrates existing credentials and retains independent endpoints for one provider', () => {
    const legacy = resolveAiSettings(
      { baseUrl: 'http://localhost:1234/v1', apiKey: 'existing', model: 'old-model' },
      defaultAiSettings(),
    )
    expect(legacy.modelProfiles?.[0]?.config.apiKey).toBe('existing')
    const alternate = saveModelProfile(
      {
        ...legacy,
        providers: {
          ...legacy.providers,
          custom: {
            apiKey: 'other',
            model: 'new-model',
            baseUrl: 'http://localhost:5678/v1',
          },
        },
      },
      'other',
      'Other endpoint',
    )
    const original = activateModelProfile(alternate, 'default')
    expect(original.providers.custom.model).toBe('old-model')
    expect(activateModelProfile(original, 'other').providers.custom.baseUrl).toBe(
      'http://localhost:5678/v1',
    )
    expect(removeModelProfile(alternate, 'other').activeModelProfileId).toBe('default')
    expect(() => removeModelProfile(original, 'missing')).not.toThrow()
    expect(() => removeModelProfile(legacy, 'default')).toThrow(/at least one/)
  })

  it('persists and broadcasts saves and profile changes to already open chats', () => {
    let stored: AiSettings = resolveAiSettings({}, defaultAiSettings())
    const handlers = new Map<string, (...args: any[]) => any>()
    const events: AiSettings[] = []
    registerAiSettingsHandlers({
      ipc: {
        handle: (channel, handler) => {
          handlers.set(channel, handler)
        },
      },
      read: () => stored,
      write: (settings) => {
        stored = JSON.parse(JSON.stringify(settings))
      },
      broadcast: (settings) => {
        events.push(settings)
      },
    })
    const input = {
      ...stored,
      provider: 'custom' as const,
      providers: {
        ...stored.providers,
        custom: { apiKey: 'test-only', model: 'model-b', baseUrl: 'http://localhost:1234/v1' },
      },
    }
    handlers.get('ai:save-model-profile')!(null, input, 'b', 'Work')
    expect(events.at(-1)?.providers.custom.model).toBe('model-b')
    handlers.get('ai:activate-model-profile')!(null, 'default')
    expect(stored.activeModelProfileId).toBe('default')
    handlers.get('ai:activate-model-profile')!(null, 'b')
    expect(stored.providers.custom.model).toBe('model-b')
    expect(handlers.get('ai:get-settings')!().modelProfiles).toHaveLength(2)
    expect(events).toHaveLength(3)
    handlers.get('ai:set-settings')!(null, {
      ...stored,
      providers: {
        ...stored.providers,
        custom: { ...stored.providers.custom, model: 'edited-model' },
      },
    })
    expect(events.at(-1)?.providers.custom.model).toBe('edited-model')
    expect(stored.modelProfiles?.find((p) => p.id === 'b')?.config.model).toBe('edited-model')
  })
})
