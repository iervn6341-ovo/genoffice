import type { AiModelProfile, AiSettings } from './types'

/** Upgrade existing single-model settings without losing their credentials. */
export function withModelProfiles(settings: AiSettings): AiSettings {
  const seen = new Set<string>()
  const profiles = (Array.isArray(settings.modelProfiles) ? settings.modelProfiles : [])
    .filter((profile): profile is AiModelProfile => {
      if (
        !profile ||
        typeof profile.id !== 'string' ||
        !profile.id ||
        seen.has(profile.id) ||
        typeof profile.name !== 'string' ||
        !profile.name.trim() ||
        !Object.hasOwn(settings.providers, profile.provider) ||
        typeof profile.config?.apiKey !== 'string' ||
        typeof profile.config?.model !== 'string'
      )
        return false
      seen.add(profile.id)
      return true
    })
    .slice(0, 100)
    .map((profile) => ({
      ...profile,
      name: profile.name.trim().slice(0, 80),
      config: { ...profile.config },
    }))
  if (!profiles.length)
    profiles.push({
      id: 'default',
      name: 'Default',
      provider: settings.provider,
      config: { ...settings.providers[settings.provider] },
      maxOutputTokens: settings.maxOutputTokens,
    })
  const activeModelProfileId = profiles.some((p) => p.id === settings.activeModelProfileId)
    ? settings.activeModelProfileId!
    : profiles[0]!.id
  return { ...settings, modelProfiles: profiles, activeModelProfileId }
}

export function saveModelProfile(settings: AiSettings, id: string, name: string): AiSettings {
  const normalized = withModelProfiles(settings)
  if (!id || !name.trim()) throw new Error('A profile name is required')
  const profile: AiModelProfile = {
    id,
    name: name.trim().slice(0, 80),
    provider: settings.provider,
    config: { ...settings.providers[settings.provider] },
    maxOutputTokens: settings.maxOutputTokens,
  }
  const profiles = normalized.modelProfiles!.filter((p) => p.id !== id)
  if (profiles.length >= 100) throw new Error('At most 100 model profiles can be saved')
  const oldIndex = normalized.modelProfiles!.findIndex((p) => p.id === id)
  profiles.splice(oldIndex < 0 ? profiles.length : oldIndex, 0, profile)
  return { ...normalized, modelProfiles: profiles, activeModelProfileId: id }
}

export function activateModelProfile(settings: AiSettings, id: string): AiSettings {
  const normalized = withModelProfiles(settings)
  const profile = normalized.modelProfiles!.find((p) => p.id === id)
  if (!profile) throw new Error('Model profile not found')
  return {
    ...normalized,
    activeModelProfileId: id,
    provider: profile.provider,
    providers: { ...normalized.providers, [profile.provider]: { ...profile.config } },
    maxOutputTokens: profile.maxOutputTokens,
  }
}

export function removeModelProfile(settings: AiSettings, id: string): AiSettings {
  const normalized = withModelProfiles(settings)
  const profiles = normalized.modelProfiles!.filter((p) => p.id !== id)
  if (!profiles.length) throw new Error('Keep at least one model profile')
  const next = { ...normalized, modelProfiles: profiles }
  return normalized.activeModelProfileId === id ? activateModelProfile(next, profiles[0]!.id) : next
}
