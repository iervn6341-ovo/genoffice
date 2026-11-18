import { expect, test, type Page } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

type Theme = 'light' | 'dark' | 'gray' | 'system'
const theme = (page: Page) =>
  page.evaluate(() => ({
    family: document.documentElement.getAttribute('data-theme'),
    variant: document.documentElement.getAttribute('data-theme-variant'),
  }))
const choose = (page: Page, value: Theme) =>
  page.evaluate((t) => window.aiOffice.setTheme(t), value)
const appearance = (value: Theme) => ({
  family: value === 'system' ? null : value === 'gray' ? 'dark' : value,
  variant: value === 'gray' ? 'gray' : null,
})

// Reproduce a theme change after the new window read its initial snapshot but
// before another bootstrap request completes. Both windows share one main process.
test('a new shell keeps theme changes received during bootstrap', async () => {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-theme-bootstrap-'))
  const file = join(root, 'theme-race.md')
  await writeFile(file, '# Theme race\n')
  const l = await launchShell({ onboardingSeen: true, openFile: file, videoDir: 'theme-bootstrap' })
  try {
    const shell = await waitForPageWithUrl(l.app, '/shell/out/renderer/')
    await choose(shell, 'dark')
    await expect.poll(() => theme(shell)).toEqual(appearance('dark'))
    const tab = (await shell.evaluate(() => window.aiOfficeTabs.list())).find(
      (t) => t.kind === 'markdown',
    )!
    const bounds = await l.app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()[0].getContentBounds(),
    )
    await l.app.evaluate(({ ipcMain }) => {
      const state = globalThis as any
      state.__themeSnapshotRead = false
      ipcMain.removeHandler('home:get-language')
      ipcMain.handle(
        'home:get-language',
        () =>
          new Promise<string>((resolve) => {
            state.__releaseThemeBootstrap = () => resolve('en')
          }),
      )
      ipcMain.removeHandler('home:get-theme')
      ipcMain.handle('home:get-theme', () => {
        state.__themeSnapshotRead = true
        return 'dark'
      })
    })
    const born = l.app.waitForEvent('window')
    const moving = shell.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
      id: tab.id,
      x: bounds.x + 250,
      y: bounds.y + 250,
    })
    await expect
      .poll(() =>
        l.app.evaluate(() =>
          Boolean(
            (globalThis as any).__themeSnapshotRead && (globalThis as any).__releaseThemeBootstrap,
          ),
        ),
      )
      .toBe(true)
    await choose(shell, 'light')
    await expect.poll(() => theme(shell)).toEqual(appearance('light'))
    await l.app.evaluate(() => (globalThis as any).__releaseThemeBootstrap())
    expect(await moving).toBe(true)
    const detached = await born
    await expect.poll(() => theme(detached)).toEqual(appearance('light'))
    const editor = await waitForPageWithUrl(l.app, '://markdown/')
    await expect.poll(() => theme(editor)).toEqual(appearance('light'))
  } finally {
    await l.app.evaluate(() => (globalThis as any).__releaseThemeBootstrap?.()).catch(() => {})
    await closeAndSaveVideo(l, 'theme-bootstrap')
  }
})

test('themes stay synchronized across detach, both setting windows, merge and reopen', async () => {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-theme-windows-'))
  const file = join(root, 'theme-window.md')
  const authored = '# Authored document\n\nTheme changes preserve this content.\n'
  await writeFile(file, authored)
  const l = await launchShell({ onboardingSeen: true, openFile: file, videoDir: 'theme-windows' })
  try {
    const shell = await waitForPageWithUrl(l.app, '/shell/out/renderer/')
    const editor = await waitForPageWithUrl(l.app, '://markdown/')
    await expect(editor.locator('.doc-editor')).toContainText('Authored document')
    await choose(shell, 'gray')
    const tab = (await shell.evaluate(() => window.aiOfficeTabs.list())).find(
      (t) => t.kind === 'markdown',
    )!
    const origin = await l.app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0]
      win.setBounds({ x: 50, y: 60, width: 940, height: 650 })
      return { id: win.id, ...win.getContentBounds() }
    })
    const born = l.app.waitForEvent('window')
    expect(
      await shell.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
        id: tab.id,
        x: origin.x + 250,
        y: origin.y + 250,
      }),
    ).toBe(true)
    const detached = await born
    for (const page of [shell, detached, editor])
      await expect.poll(() => theme(page)).toEqual(appearance('gray'))
    for (const page of [shell, detached]) {
      await page.evaluate(() => window.aiOfficeTabs.activate('home'))
      await page.getByRole('button', { name: 'Settings', exact: true }).click()
      await page.getByRole('button', { name: 'General', exact: true }).click()
      await expect(page.getByRole('button', { name: 'Theme', exact: true })).toHaveAttribute(
        'data-value',
        'gray',
      )
    }
    // An already-open settings panel must reflect a choice made elsewhere.
    await detached.getByRole('button', { name: 'Theme', exact: true }).click()
    await detached.getByRole('option', { name: 'Light', exact: true }).click()
    for (const page of [shell, detached])
      await expect(page.getByRole('button', { name: 'Theme', exact: true })).toHaveAttribute(
        'data-value',
        'light',
      )
    for (const [sender, next] of [
      [detached, 'light'],
      [shell, 'dark'],
      [detached, 'system'],
    ] as const) {
      await choose(sender, next)
      for (const page of [shell, detached, editor])
        await expect.poll(() => theme(page)).toEqual(appearance(next))
      for (const page of [shell, detached])
        await expect(page.getByRole('button', { name: 'Theme', exact: true })).toHaveAttribute(
          'data-value',
          next,
        )
    }
    // Simulate the system color preference within this test process only.
    // Playwright defaults to emulated light, which would mask native changes.
    for (const page of [shell, detached, editor]) await page.emulateMedia({ colorScheme: null })
    for (const scheme of ['dark', 'light'] as const) {
      await l.app.evaluate(({ nativeTheme }, scheme) => {
        nativeTheme.themeSource = scheme
      }, scheme)
      for (const page of [shell, detached, editor])
        await expect
          .poll(() => page.evaluate(() => matchMedia('(prefers-color-scheme: dark)').matches))
          .toBe(scheme === 'dark')
    }
    await choose(detached, 'gray')
    for (const page of [shell, detached]) await page.locator('.set-close').click()
    await l.app.evaluate(({ BrowserWindow }, id) => {
      const target = BrowserWindow.getAllWindows().find((w) => w.id !== id)!
      target.setPosition(350, 300)
    }, origin.id)
    expect(
      await detached.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
        id: tab.id,
        x: origin.x + 200,
        y: origin.y + 20,
      }),
    ).toBe(true)
    await expect.poll(() => detached.isClosed()).toBe(true)
    for (const page of [shell, editor])
      await expect.poll(() => theme(page)).toEqual(appearance('gray'))
    expect(await readFile(file, 'utf8')).toBe(authored)
  } finally {
    await closeAndSaveVideo(l, 'theme-windows')
  }
  const again = await launchShell({
    userDataDir: l.userDataDir,
    openFile: file,
    videoDir: 'theme-windows-reopen',
  })
  try {
    const shell = await waitForPageWithUrl(again.app, '/shell/out/renderer/')
    const editor = await waitForPageWithUrl(again.app, '://markdown/')
    for (const page of [shell, editor])
      await expect.poll(() => theme(page)).toEqual(appearance('gray'))
    expect(await readFile(file, 'utf8')).toBe(authored)
  } finally {
    await closeAndSaveVideo(again, 'theme-windows-reopen')
  }
})
