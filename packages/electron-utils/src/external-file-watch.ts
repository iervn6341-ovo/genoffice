import { watch, type FSWatcher } from 'node:fs'
import { basename, dirname } from 'node:path'

/** Watch the directory so Office's atomic replace saves keep being detected. */
export class ExternalFileWatch {
  private handles = new Map<number, { path: string; watcher: FSWatcher; timer?: NodeJS.Timeout }>()

  remember(id: number, path: string, changed: () => void): void {
    if (this.handles.get(id)?.path === path) return
    this.forget(id)
    try {
      const watcher = watch(dirname(path), { persistent: false }, (_event, filename) => {
        if (filename != null && filename.toString() !== basename(path)) return
        const entry = this.handles.get(id)
        if (!entry || entry.watcher !== watcher) return
        if (entry.timer) clearTimeout(entry.timer)
        entry.timer = setTimeout(() => {
          entry.timer = undefined
          changed()
        }, 250)
        entry.timer.unref()
      })
      this.handles.set(id, { path, watcher })
      // Focus checks still work on volumes that do not support notifications.
      watcher.on('error', () => {
        if (this.handles.get(id)?.watcher === watcher) this.forget(id)
      })
    } catch {
      /* Check on focus/save instead; never fall back to polling. */
    }
  }

  forget(id: number): void {
    const entry = this.handles.get(id)
    if (entry?.timer) clearTimeout(entry.timer)
    entry?.watcher.close()
    this.handles.delete(id)
  }
}
