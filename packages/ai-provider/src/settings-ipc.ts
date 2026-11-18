import { defaultAiSettings, resolveAiSettings } from './providers'
import { activateModelProfile, removeModelProfile, saveModelProfile } from './model-profiles'
import type { AiSettings } from './types'

/** Shared by standalone editors and the shell so every settings change reaches open chats. */
export function registerAiSettingsHandlers(options: {
  ipc: { handle(channel: string, listener: (...args: any[]) => any): void }
  read(): Partial<AiSettings>
  write(settings: AiSettings): void
  broadcast(settings: AiSettings): void
}): void {
  const read = () => resolveAiSettings(options.read(), defaultAiSettings())
  const commit = (settings: AiSettings) => {
    options.write(settings)
    options.broadcast(settings)
    return settings
  }
  options.ipc.handle('ai:get-settings', read)
  options.ipc.handle('ai:set-settings', (_event, input: AiSettings) => {
    const current = read()
    const next = resolveAiSettings(
      {
        ...input,
        modelProfiles: current.modelProfiles,
        activeModelProfileId: current.activeModelProfileId,
      },
      defaultAiSettings(),
    )
    const active = current.modelProfiles!.find((p) => p.id === current.activeModelProfileId)!
    return commit(saveModelProfile(next, active.id, active.name))
  })
  options.ipc.handle(
    'ai:save-model-profile',
    (_event, input: AiSettings, id: string, name: string) => {
      const current = read()
      const next = resolveAiSettings(
        {
          ...current,
          provider: input.provider,
          providers: { ...current.providers, [input.provider]: input.providers[input.provider] },
          maxOutputTokens: input.maxOutputTokens,
        },
        defaultAiSettings(),
      )
      if (typeof id !== 'string' || typeof name !== 'string')
        throw new Error('Invalid model profile')
      return commit(saveModelProfile(next, id, name))
    },
  )
  options.ipc.handle('ai:activate-model-profile', (_event, id: string) =>
    commit(activateModelProfile(read(), id)),
  )
  options.ipc.handle('ai:remove-model-profile', (_event, id: string) =>
    commit(removeModelProfile(read(), id)),
  )
}
