import { contextBridge, ipcRenderer } from 'electron'

/** One bridge shared by the settings page and all six chat panels. */
export function installAiSettingsBridge(): void {
  contextBridge.exposeInMainWorld('aiModelSettings', {
    get: () => ipcRenderer.invoke('ai:get-settings'),
    saveProfile: (settings: unknown, id: string, name: string) =>
      ipcRenderer.invoke('ai:save-model-profile', settings, id, name),
    activate: (id: string) => ipcRenderer.invoke('ai:activate-model-profile', id),
    remove: (id: string) => ipcRenderer.invoke('ai:remove-model-profile', id),
    onChanged: (handler: (settings: unknown) => void) => {
      const listener = (_event: unknown, settings: unknown) => handler(settings)
      ipcRenderer.on('ai:settings-changed', listener)
      return () => ipcRenderer.removeListener('ai:settings-changed', listener)
    },
  })
}
