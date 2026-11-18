import { expect, test } from '@playwright/test'
import { launchShell, closeAndSaveVideo } from './helpers'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { openFreshEditor } from './editor-lifecycle-helpers'

test('Home menus preserve native title-bar dragging and dismiss on move or resize', async () => {
  const l = await launchShell({ onboardingSeen: true, videoDir: 'native-window-chrome' })
  try {
    const bar = l.page.locator('.tab-bar')
    const menuTrigger = l.page.getByRole('button', {
      name: 'More actions for GenOffice',
      exact: true,
    })
    const menu = l.page.locator('.folder-menu')
    for (const event of ['will-move', 'will-resize'] as const) {
      // Let an earlier native gesture finish; the relay only dismisses once per burst.
      await l.page.waitForTimeout(170)
      await menuTrigger.click()
      await expect(menu).toBeVisible()
      await expect(l.page.locator('html')).toHaveClass(/genoffice-popover-open/)
      expect(
        await bar.evaluate((el) => getComputedStyle(el).getPropertyValue('-webkit-app-region')),
      ).toBe('drag')
      expect(
        await l.page
          .locator('.tab-home')
          .evaluate((el) => getComputedStyle(el).getPropertyValue('-webkit-app-region')),
      ).toBe('no-drag')
      const result = await l.app.evaluate(({ BrowserWindow }, event) => {
        const w = BrowserWindow.getAllWindows()[0]
        const before = w.getBounds()
        let canceled = false
        w.emit(
          event,
          {
            preventDefault: () => {
              canceled = true
            },
          },
          before,
        )
        return { before, after: w.getBounds(), canceled }
      }, event)
      expect(result.after).toEqual(result.before)
      expect(result.canceled).toBe(false)
      await expect(menu).toBeHidden()
      await expect(l.page.locator('html')).not.toHaveClass(/genoffice-popover-open/)
    }
  } finally {
    await closeAndSaveVideo(l, 'native-window-chrome')
  }
})

test('drag previews name the destination, mark its slot and disappear on cancellation', async () => {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-drag-preview-'))
  const md = join(root, 'target.md')
  const html = join(root, 'source.html')
  await writeFile(md, '# Target preview\n')
  await writeFile(html, '<!doctype html><p>Source preview</p>')
  const l = await launchShell({ onboardingSeen: true, videoDir: 'tab-drag-preview' })
  try {
    await openFreshEditor(l, md, 'markdown')
    let tab = (await l.page.evaluate(() => window.aiOfficeTabs.list())).find(
      (t) => t.kind === 'markdown',
    )!
    const firstBounds = await l.app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()[0].getContentBounds(),
    )
    const born = l.app.waitForEvent('window')
    expect(
      await l.page.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
        id: tab.id,
        x: firstBounds.x + 250,
        y: firstBounds.y + 200,
      }),
    ).toBe(true)
    const second = await born
    await second.waitForFunction(() => !!window.aiOfficeTabs)
    if (process.platform === 'darwin') {
      const native = await l.app.evaluate(({ BrowserWindow, app }) => {
        const version = process.getSystemVersion()
        if (!version.startsWith('27.2')) return { applicable: false }
        const require = process
          .getBuiltinModule('module')
          .createRequire(app.getAppPath() + '/package.json')
        const path = process.getBuiltinModule('path')
        const guard = require(path.join(app.getAppPath(), 'out/native/window-gesture-guard.node'))
        return {
          applicable: true,
          // Both the original Home shell and the detached target must recover.
          windows: BrowserWindow.getAllWindows().map((w) => ({
            id: w.id,
            ...guard.stats(w.getNativeWindowHandle()),
          })),
        }
      })
      if (native.applicable) {
        expect(native.windows).toHaveLength(2)
        expect(native.windows.every((w) => w.installed && !w.active)).toBe(true)
      }
    }
    await openFreshEditor(l, html, 'html')
    tab = (await l.page.evaluate(() => window.aiOfficeTabs.list())).find((t) => t.kind === 'html')!
    const targetBounds = await l.app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()
        .find((w) => w.getTitle().startsWith('target.md'))!
        .getContentBounds(),
    )
    const previewBorn = l.app.waitForEvent('window')
    await l.page.evaluate(({ id, x, y }) => window.aiOfficeTabs.previewDrag(id, x, y), {
      id: tab.id,
      x: targetBounds.x + 400,
      y: targetBounds.y + 20,
    })
    const preview = await previewBorn
    await expect(preview.locator('#title')).toHaveText('source.html')
    await expect(preview.locator('#label')).toContainText('target.md')
    await expect(preview.locator('img')).toHaveAttribute('src', /^data:image/)
    await expect(second.locator('.tab-drop-marker')).toBeVisible()
    expect(
      await l.app.evaluate(({ BrowserWindow }) => {
        const w = BrowserWindow.getAllWindows().find(
          (w) => w.getTitle() === 'GenOffice Tab Preview',
        )!
        return { focusable: w.isFocusable(), focused: w.isFocused() }
      }),
    ).toEqual({ focusable: false, focused: false })
    await l.page.evaluate(() => window.aiOfficeTabs.endPreview())
    await expect.poll(() => preview.isClosed()).toBe(true)
    await expect(second.locator('.tab-drop-marker')).toHaveCount(0)
    expect(
      (await l.page.evaluate(() => window.aiOfficeTabs.list())).some((t) => t.kind === 'html'),
    ).toBe(true)
    expect(
      (await second.evaluate(() => window.aiOfficeTabs.list())).some((t) => t.kind === 'markdown'),
    ).toBe(true)

    const nextPreview = l.app.waitForEvent('window')
    await l.page.evaluate(({ id, x, y }) => window.aiOfficeTabs.previewDrag(id, x, y), {
      id: tab.id,
      x: targetBounds.x + 400,
      y: targetBounds.y + 200,
    })
    const detachedPreview = await nextPreview
    await expect(detachedPreview.locator('#label')).toHaveText('Release to open a new window')
    await l.page.evaluate(() => window.aiOfficeTabs.endPreview())
    await expect.poll(() => detachedPreview.isClosed()).toBe(true)
    expect(await l.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(
      2,
    )

    // Return the live tab to the original shell: this was the unguarded path.
    const originalBounds = await l.app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()
        .find((w) => w.getTitle().startsWith('source.html'))!
        .getContentBounds(),
    )
    const returning = (await second.evaluate(() => window.aiOfficeTabs.list())).find(
      (t) => t.kind === 'markdown',
    )!
    expect(
      await second.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
        id: returning.id,
        x: originalBounds.x + 300,
        y: originalBounds.y + 20,
      }),
    ).toBe(true)
    await expect.poll(() => second.isClosed()).toBe(true)
    expect(
      (await l.page.evaluate(() => window.aiOfficeTabs.list())).some((t) => t.id === returning.id),
    ).toBe(true)
    if (process.platform === 'darwin') {
      const installed = await l.app.evaluate(({ BrowserWindow, app }) => {
        if (!process.getSystemVersion().startsWith('27.2')) return true
        const require = process
          .getBuiltinModule('module')
          .createRequire(app.getAppPath() + '/package.json')
        const guard = require(app.getAppPath() + '/out/native/window-gesture-guard.node')
        return guard.stats(BrowserWindow.getAllWindows()[0].getNativeWindowHandle()).installed
      })
      expect(installed).toBe(true)
    }
  } finally {
    await closeAndSaveVideo(l, 'tab-drag-preview')
  }
})
