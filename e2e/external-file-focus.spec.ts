import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile, rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('external changes use OS events and wait for the foreground document', async () => {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-external-focus-'))
  const file = join(root, 'focus.md')
  await writeFile(file, '# Original\n')
  const l = await launchShell({
    onboardingSeen: true,
    foreground: true,
    openFile: file,
    videoDir: 'external-file-focus',
    settings: { autoSave: false },
  })
  try {
    // Opening a file at launch can attach its editor before the shell page.
    // The tabs API belongs to the shell, regardless of firstWindow() order.
    const shell = await waitForPageWithUrl(l.app, '/shell/out/renderer/')
    const page = await waitForPageWithUrl(l.app, '://markdown/')
    await expect(page.locator('.doc-editor')).toContainText('Original')
    await l.app.evaluate(({ dialog }) => {
      ;(globalThis as any).__externalNotices = 0
      dialog.showMessageBox = (async (...args: any[]) => {
        const opts = args.at(-1)
        if (opts.title === 'File changed outside GenOffice') (globalThis as any).__externalNotices++
        return { response: opts.cancelId ?? 1, checkboxChecked: false }
      }) as any
    })
    const calls = () => l.app.evaluate(() => (globalThis as any).__externalNotices)
    await l.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].hide())
    await writeFile(file, '# Changed while app hidden\n')
    await page.waitForTimeout(700)
    expect(await calls()).toBe(0)
    const focus = () =>
      l.app.evaluate(({ app, BrowserWindow, webContents }) => {
        const win = BrowserWindow.getAllWindows()[0]
        app.focus({ steal: true })
        win.show()
        win.focus()
        webContents
          .getAllWebContents()
          .find((w) => w.getURL().includes('://markdown/'))!
          .focus()
      })
    await focus()
    await expect.poll(calls).toBe(1)
    await writeFile(join(root, 'replacement.tmp'), '# Atomic Office-style replacement\n')
    await rename(join(root, 'replacement.tmp'), file)
    await expect.poll(calls).toBe(2)
    const tab = (await shell.evaluate(() => window.aiOfficeTabs.list())).find(
      (t) => t.kind === 'markdown',
    )!
    await shell.evaluate(() => window.aiOfficeTabs.activate('home'))
    // Simulate a native first responder retained by a hidden document view.
    await l.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((w) => w.getURL().includes('://markdown/'))!
        .focus()
    })
    await writeFile(file, '# Changed in background tab\n')
    await page.waitForTimeout(700)
    expect(await calls()).toBe(2)
    await shell.evaluate((id) => window.aiOfficeTabs.activate(id), tab.id)
    await focus()
    await expect.poll(calls).toBe(3)
  } finally {
    await closeAndSaveVideo(l, 'external-file-focus')
  }
})
