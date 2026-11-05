import { useEffect, useRef, useState } from 'react'
import type { AiSettings } from '@genoffice/ai-provider'
import './ai-model-profiles.css'

export interface AiModelSettingsBridge {
  get(): Promise<AiSettings>
  saveProfile(settings: AiSettings, id: string, name: string): Promise<AiSettings>
  activate(id: string): Promise<AiSettings>
  remove(id: string): Promise<AiSettings>
  onChanged(handler: (settings: AiSettings) => void): () => void
}

declare global {
  interface Window {
    aiModelSettings?: AiModelSettingsBridge
  }
}

export function aiProfileStrings(lang: string) {
  if (lang === 'zh-TW')
    return {
      profiles: '模型設定檔',
      add: '新增設定檔',
      name: '設定檔名稱',
      save: '儲存並使用',
      use: '使用此設定檔',
      remove: '刪除設定檔',
      active: '使用中',
      default: '預設設定',
      draft: '新設定檔',
      hint: '測試連線後，請按「儲存並使用」套用至聊天。',
      error: '無法更新模型設定檔',
      discard: '捨棄尚未儲存的設定？',
      confirmRemove: '刪除此模型設定檔？',
    }
  if (lang === 'zh' || lang === 'zh-CN')
    return {
      profiles: '模型配置',
      add: '新增配置',
      name: '配置名称',
      save: '保存并使用',
      use: '使用此配置',
      remove: '删除配置',
      active: '使用中',
      default: '默认配置',
      draft: '新配置',
      hint: '测试连接后，请按“保存并使用”应用到聊天。',
      error: '无法更新模型配置',
      discard: '放弃尚未保存的设置？',
      confirmRemove: '删除此模型配置？',
    }
  return {
    profiles: 'Model profiles',
    add: 'New profile',
    name: 'Profile name',
    save: 'Save and use',
    use: 'Use this profile',
    remove: 'Delete profile',
    active: 'Active',
    default: 'Default',
    draft: 'New profile',
    hint: 'After testing the connection, choose “Save and use” to apply it to chat.',
    error: 'Could not update model profile',
    discard: 'Discard unsaved settings?',
    confirmRemove: 'Delete this model profile?',
  }
}

/** Subscribe before the initial read so a late read cannot undo a saved setting. */
export function useAiSettingsSync(onSettings: (settings: AiSettings) => void): void {
  const callback = useRef(onSettings)
  callback.current = onSettings
  useEffect(() => {
    const bridge = window.aiModelSettings
    if (!bridge) return
    let alive = true
    let revision = 0
    const unsubscribe = bridge.onChanged((settings) => {
      revision++
      if (alive) callback.current(settings)
    })
    void bridge
      .get()
      .then((settings) => {
        if (alive && revision === 0) callback.current(settings)
      })
      .catch(() => {})
    return () => {
      alive = false
      unsubscribe()
    }
  }, [])
}

export function AiModelProfilePicker({
  lang,
  disabled = false,
}: {
  lang: string
  disabled?: boolean
}) {
  const [settings, setSettings] = useState<AiSettings | null>(null)
  const [pending, setPending] = useState(false)
  const [error, setError] = useState('')
  useAiSettingsSync(setSettings)
  const labels = aiProfileStrings(lang)
  if (!settings?.modelProfiles?.length) return null
  return (
    <div className="ai-model-profile-picker">
      <select
        aria-label={labels.profiles}
        title={labels.profiles}
        value={settings.activeModelProfileId}
        disabled={disabled || pending}
        onChange={(event) => {
          setPending(true)
          setError('')
          void window
            .aiModelSettings!.activate(event.target.value)
            .then(setSettings)
            .catch((err) => setError(err instanceof Error ? err.message : labels.error))
            .finally(() => setPending(false))
        }}
      >
        {settings.modelProfiles.map((profile) => (
          <option key={profile.id} value={profile.id}>
            {profile.id === 'default' && profile.name === 'Default' ? labels.default : profile.name}
          </option>
        ))}
      </select>
      {error && (
        <span role="alert" className="ai-model-profile-error">
          {error}
        </span>
      )}
    </div>
  )
}
