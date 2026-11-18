import { app, BrowserWindow, dialog } from 'electron'
import { ExternalFileWatch } from './external-file-watch'
import type { WebContents } from 'electron'
import { basename } from 'node:path'
import { stat } from 'node:fs/promises'
import {
  diskContentHash,
  ExternalFileChangedError,
  ExternalFileTracker,
  fileContentHash,
} from './external-file-tracker'

export { ExternalFileChangedError, fileContentHash }

export interface ExternalConflictInfo {
  path: string
  fileName: string
  missing: boolean
  mtimeMs?: number | undefined
  size?: number | undefined
}

const tracker = new ExternalFileTracker()
interface Owner {
  contents: WebContents
  language: () => string
  isDirty: () => boolean | Promise<boolean>
  reload: (path: string) => void | Promise<void>
  saveCopy: () => void | Promise<unknown>
  onConflict?: (conflict: ExternalConflictInfo) => boolean | void | Promise<boolean | void>
}
const owners = new Map<number, Owner>()
const watcher = new ExternalFileWatch()
const pendingNotices = new Map<number, { path: string; missing: boolean }>()
let notices: Promise<unknown> = Promise.resolve()
let foregroundTabCheck: (id: number) => boolean = () => true
let backgroundIdleDelayMs = 2500
const backgroundTimers = new Map<number, NodeJS.Timeout>()

function isOwnerDestroyed(owner: Owner): boolean {
  return typeof owner.contents.isDestroyed === 'function' ? owner.contents.isDestroyed() : false
}

function focused(owner: Owner): boolean {
  if (isOwnerDestroyed(owner)) return false
  if (
    typeof owner.contents.isFocused === 'function' &&
    !owner.contents.isFocused()
  ) {
    return false
  }
  const fromWebContents = BrowserWindow?.fromWebContents
  const parent = typeof fromWebContents === 'function' ? fromWebContents(owner.contents) : null
  const getFocusedWindow = BrowserWindow?.getFocusedWindow
  const focusedWin = typeof getFocusedWindow === 'function' ? getFocusedWindow() : null
  return (
    !!parent &&
    parent === focusedWin &&
    Boolean(parent.isFocused?.()) &&
    foregroundTabCheck(owner.contents.id)
  )
}

function clearBackgroundTimer(id: number): void {
  const timer = backgroundTimers.get(id)
  if (timer) {
    clearTimeout(timer)
    backgroundTimers.delete(id)
  }
}

function scheduleBackgroundIdleCheck(owner: Owner, delay = backgroundIdleDelayMs): void {
  const id = owner.contents.id
  clearBackgroundTimer(id)
  const timer = setTimeout(async () => {
    backgroundTimers.delete(id)
    if (isOwnerDestroyed(owner)) return
    await checkBackground(owner)
  }, delay)
  timer.unref()
  backgroundTimers.set(id, timer)
}

async function checkBackground(owner: Owner): Promise<void> {
  if (isOwnerDestroyed(owner)) return
  const id = owner.contents.id
  const path = tracker.path(id)
  if (!path) return
  watcher.remember(id, path, () => onFileWatcherChange(owner))
  await tracker.check(id).catch(() => {})
}

function onFileWatcherChange(owner: Owner): void {
  if (focused(owner)) {
    clearBackgroundTimer(owner.contents.id)
    checkFocused(owner)
  } else {
    scheduleBackgroundIdleCheck(owner)
  }
}

function checkFocused(owner: Owner): void {
  clearBackgroundTimer(owner.contents.id)
  if (!focused(owner)) {
    scheduleBackgroundIdleCheck(owner)
    return
  }
  const id = owner.contents.id
  const path = tracker.path(id)
  if (path) watcher.remember(id, path, () => onFileWatcherChange(owner))
  const pending = pendingNotices.get(id)
  if (pending) void notify(owner, pending.path, pending.missing)
  void tracker.check(id).catch(() => {})
}

app.on('browser-window-focus', () => {
  for (const owner of owners.values()) checkFocused(owner)
})

app.on('browser-window-blur', () => {
  for (const owner of owners.values()) scheduleBackgroundIdleCheck(owner)
})

export function externalFileLabels(lang: string) {
  if (lang === 'zh-TW')
    return {
      title: '檔案已在外部變更',
      changed: '已被其他程式更新。',
      missing: '已在外部移動或刪除。',
      missingDetail: '目前內容仍保留在編輯器中，可另存副本。',
      clean: '是否重新載入磁碟上的最新版本？',
      dirty: '目前有尚未儲存的編輯。保留編輯不會覆寫外部版本；可另存副本。重新載入會放棄目前編輯。',
      keep: '保留編輯',
      later: '稍後',
      copy: '另存副本…',
      reload: '重新載入',
      failed: '無法重新載入檔案，已保留目前編輯。',
      blocked: '檔案已在外部變更。請保留編輯、另存副本，或重新載入檔案。',
    }
  if (lang === 'zh')
    return {
      title: '文件已在外部更改',
      changed: '已被其他程序更新。',
      missing: '已在外部移动或删除。',
      missingDetail: '当前内容仍保留在编辑器中，可另存副本。',
      clean: '是否重新加载磁盘上的最新版本？',
      dirty: '当前有尚未保存的编辑。保留编辑不会覆盖外部版本；可另存副本。重新加载会放弃当前编辑。',
      keep: '保留编辑',
      later: '稍后',
      copy: '另存副本…',
      reload: '重新加载',
      failed: '无法重新加载文件，已保留当前编辑。',
      blocked: '文件已在外部更改。请保留编辑、另存副本，或重新加载文件。',
    }
  return {
    title: 'File changed outside GenOffice',
    changed: 'was updated by another program.',
    missing: 'was moved or deleted outside GenOffice.',
    missingDetail: 'Your current content is still in the editor. You can save a copy.',
    clean: 'Reload the latest version from disk?',
    dirty:
      'You have unsaved edits. Keep them without overwriting the external version, or save a copy. Reloading discards your current edits.',
    keep: 'Keep Editing',
    later: 'Later',
    copy: 'Save a Copy…',
    reload: 'Reload',
    failed: 'Could not reload the file. Your current edits have been kept.',
    blocked:
      'The file changed outside GenOffice. Keep your edits, save a copy, or reload the file.',
  }
}

function valid(owner: Owner, path: string) {
  return (
    owners.get(owner.contents.id) === owner &&
    !isOwnerDestroyed(owner) &&
    tracker.path(owner.contents.id) === path
  )
}

function notify(owner: Owner, path: string, missing: boolean): Promise<void> {
  const notice = { path, missing }
  pendingNotices.set(owner.contents.id, notice)
  const next = notices
    .catch(() => {})
    .then(async () => {
      if (
        pendingNotices.get(owner.contents.id) !== notice ||
        !valid(owner, path)
      )
        return
      const isFore = focused(owner)
      const dirty = await owner.isDirty()
      if (!valid(owner, path)) return

      // In background: defer prompt until the window/tab gains focus,
      // so external changes (e.g. from Claude Code or terminal CLI) always prompt
      // the user upon returning rather than stealing focus or silently auto-reloading.
      if (!isFore) {
        if (owner.onConflict && dirty) {
          const stats = await stat(path).catch(() => null)
          await owner.onConflict({
            path,
            fileName: basename(path),
            missing,
            mtimeMs: stats?.mtimeMs,
            size: stats?.size,
          })
        }
        return
      }

      // Foreground: if conflict handler provided and dirty, trigger custom conflict resolution
      if (dirty && owner.onConflict) {
        pendingNotices.delete(owner.contents.id)
        const stats = await stat(path).catch(() => null)
        const handled = await owner.onConflict({
          path,
          fileName: basename(path),
          missing,
          mtimeMs: stats?.mtimeMs,
          size: stats?.size,
        })
        if (handled !== false) return
      }

      pendingNotices.delete(owner.contents.id)
      const labels = externalFileLabels(owner.language())
      const options = {
        type: 'warning' as const,
        title: labels.title,
        message: `${basename(path)} ${missing ? labels.missing : labels.changed}`,
        detail: missing ? labels.missingDetail : dirty ? labels.dirty : labels.clean,
        buttons:
          dirty || missing
            ? [labels.keep, labels.copy, ...(!missing ? [labels.reload] : [])]
            : [labels.reload, labels.later],
        defaultId: dirty || missing ? 0 : 1,
        cancelId: dirty || missing ? 0 : 1,
        noLink: true,
      }
      const parent = BrowserWindow.fromWebContents(owner.contents)
      const { response } =
        parent && !parent.isDestroyed()
          ? await dialog.showMessageBox(parent, options)
          : await dialog.showMessageBox(options)
      if (!valid(owner, path)) return
      if ((dirty || missing) && response === 1) {
        await owner.saveCopy()
        return
      }
      if (missing || response !== (dirty ? 2 : 0)) return
      // A native modal does not stop all authoring routes (AI/IPC). Confirm if a
      // formerly clean editor became dirty while the prompt was displayed.
      if (!dirty && (await owner.isDirty())) {
        const confirm = {
          ...options,
          detail: labels.dirty,
          buttons: [labels.keep, labels.copy, labels.reload],
          defaultId: 0,
          cancelId: 0,
        }
        const choice =
          parent && !parent.isDestroyed()
            ? await dialog.showMessageBox(parent, confirm)
            : await dialog.showMessageBox(confirm)
        if (!valid(owner, path)) return
        if (choice.response === 1) {
          await owner.saveCopy()
          return
        }
        if (choice.response !== 2) return
      }
      try {
        await owner.reload(path)
      } catch {
        if (valid(owner, path))
          await dialog.showMessageBox({ type: 'error', message: labels.failed })
      }
    })
  notices = next
  return next
}

/** Attach once per WebContents, also for standalone editors. */
function attach(contents: WebContents, options: Omit<Owner, 'contents'>): void {
  const previous = owners.get(contents.id)
  if (previous) return
  owners.set(contents.id, { contents, ...options })
  if (typeof contents.on === 'function') {
    contents.on('focus', () => {
      checkFocused(owners.get(contents.id)!)
    })
    contents.on('blur', () => {
      const owner = owners.get(contents.id)
      if (owner) scheduleBackgroundIdleCheck(owner)
    })
  }
  if (typeof contents.once === 'function') {
    contents.once('destroyed', () => forget(contents.id))
  }
}

async function remember(id: number, path: string, hash?: string): Promise<void> {
  const owner = owners.get(id)
  if (!owner || isOwnerDestroyed(owner)) return
  pendingNotices.delete(id)
  await tracker.remember(id, path, hash ?? (await diskContentHash(path)), (missing) =>
    notify(owner, path, missing),
  )
  if (!valid(owner, path)) return
  watcher.remember(id, path, () => onFileWatcherChange(owner))
  checkFocused(owner)
}

function forget(id: number): void {
  clearBackgroundTimer(id)
  owners.delete(id)
  pendingNotices.delete(id)
  watcher.forget(id)
  tracker.forget(id)
}

async function forceRemember(id: number, path: string): Promise<void> {
  const owner = owners.get(id)
  if (!owner) return
  const hash = await diskContentHash(path).catch(() => 'missing')
  await tracker.remember(id, path, hash, (missing) => notify(owner, path, missing))
}

async function write<T>(id: number, path: string, writer: () => Promise<T>): Promise<T> {
  try {
    const result = await tracker.write(id, path, writer)
    if (tracker.path(id) === path) await remember(id, path)
    return result
  } catch (error) {
    if (error instanceof ExternalFileChangedError) {
      // Never wait for a Save a Copy action from inside the current save IPC.
      blocked(id)
      error.message = externalFileLabels(owners.get(id)?.language() ?? 'en').blocked
    }
    throw error
  }
}

function blocked(id: number): void {
  setTimeout(() => {
    const owner = owners.get(id)
    if (!owner || !focused(owner)) return
    void tracker.check(id, true).catch(() => {})
  }, 0)
}

async function renamed(id: number, oldPath: string, newPath: string): Promise<void> {
  if (tracker.path(id) !== oldPath) return
  await remember(id, newPath, tracker.hash(id))
}

export const externalFiles = {
  setBackgroundIdleDelay(ms: number): void {
    backgroundIdleDelayMs = ms
  },
  triggerBackgroundCheck(id: number): Promise<void> {
    const owner = owners.get(id)
    if (owner) return checkBackground(owner)
    return Promise.resolve()
  },
  forceRemember,
  setForegroundTabCheck(check: (id: number) => boolean): void {
    foregroundTabCheck = check
  },
  checkForeground(id: number): void {
    const owner = owners.get(id)
    if (owner) checkFocused(owner)
  },
  attach,
  remember,
  forget,
  renamed,
  write,
  blocked,
  assertWritable: tracker.assertWritable.bind(tracker),
}
