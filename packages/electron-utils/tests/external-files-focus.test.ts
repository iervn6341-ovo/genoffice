import { it, expect, vi } from 'vitest'
import { EventEmitter } from 'node:events'
import { mkdtemp, writeFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ExternalFileTracker } from '../src/external-file-tracker'

const fake = vi.hoisted(() => ({
  foreground: false,
  changed: undefined as undefined | (() => void),
  message: vi.fn(async () => ({ response: 1 })),
}))
vi.mock('electron', () => {
  const parent = { isFocused: () => fake.foreground, isDestroyed: () => false }
  return {
    app: { on: vi.fn() },
    BrowserWindow: {
      fromWebContents: () => parent,
      getFocusedWindow: () => (fake.foreground ? parent : null),
    },
    dialog: { showMessageBox: fake.message },
  }
})
vi.mock('../src/external-file-watch', () => ({
  ExternalFileWatch: class {
    remember(_id: number, _path: string, changed: () => void) {
      fake.changed = changed
    }
    forget() {
      fake.changed = undefined
    }
  },
}))

it('defers background notifications until focus, checks active OS events, and never polls', async () => {
  const checks = vi.spyOn(ExternalFileTracker.prototype, 'check')
  const interval = vi.spyOn(globalThis, 'setInterval')
  const { externalFiles, fileContentHash } = await import('../src/external-files')
  expect(interval).not.toHaveBeenCalled()
  interval.mockRestore()
  const root = await mkdtemp(join(tmpdir(), 'genoffice-focus-'))
  const path = join(root, 'file.txt')
  let focused = false
  let activeTab = true
  externalFiles.setForegroundTabCheck(() => activeTab)
  const contents = Object.assign(new EventEmitter(), {
    id: 71,
    isDestroyed: () => false,
    isFocused: () => focused,
  })
  try {
    await writeFile(path, 'original')
    externalFiles.attach(contents as any, {
      language: () => 'en',
      isDirty: () => false,
      reload: async () => {},
      saveCopy: async () => {},
    })
    await externalFiles.remember(71, path, fileContentHash('original'))
    await writeFile(path, 'external change')
    fake.changed!()
    externalFiles.blocked(71)
    await new Promise((r) => setTimeout(r, 50))
    expect(fake.message).not.toHaveBeenCalled()
    expect(checks).not.toHaveBeenCalled()
    fake.foreground = true
    focused = true
    contents.emit('focus')
    await expect.poll(() => fake.message.mock.calls.length).toBe(1)
    contents.emit('focus')
    await new Promise((r) => setTimeout(r, 50))
    expect(fake.message).toHaveBeenCalledTimes(1)
    await writeFile(path, 'another external change')
    fake.changed!()
    await expect.poll(() => fake.message.mock.calls.length).toBe(2)
    // A still-selected editor in a background OS window must also stay silent.
    fake.foreground = false
    await writeFile(path, 'changed in another application')
    fake.changed!()
    await new Promise((r) => setTimeout(r, 50))
    expect(fake.message).toHaveBeenCalledTimes(2)
    fake.foreground = true
    contents.emit('focus')
    await expect.poll(() => fake.message.mock.calls.length).toBe(3)
    // A hidden tab can retain a native first responder during tab changes.
    // Neither OS events nor a blocked background save may start disk checks.
    activeTab = false
    const checked = checks.mock.calls.length
    await writeFile(path, 'background tab change')
    fake.changed!()
    externalFiles.blocked(71)
    contents.emit('focus')
    await new Promise((r) => setTimeout(r, 50))
    expect(checks).toHaveBeenCalledTimes(checked)
    expect(fake.message).toHaveBeenCalledTimes(3)
    activeTab = true
    externalFiles.checkForeground(71)
    await expect.poll(() => fake.message.mock.calls.length).toBe(4)
    // Registering a missing background document must also wait for focus.
    fake.foreground = false
    const beforeMissing = checks.mock.calls.length
    await externalFiles.remember(71, join(root, 'missing.txt'), fileContentHash('opened'))
    await new Promise((r) => setTimeout(r, 50))
    expect(checks).toHaveBeenCalledTimes(beforeMissing)
    fake.foreground = true
    contents.emit('focus')
    await expect.poll(() => fake.message.mock.calls.length).toBe(5)
  } finally {
    externalFiles.forget(71)
    externalFiles.setForegroundTabCheck(() => true)
    checks.mockRestore()
    await rm(root, { recursive: true, force: true })
  }
})
