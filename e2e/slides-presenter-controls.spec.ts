import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

async function enterPresenter(page: Page): Promise<void> {
  await expect(page.locator('.notes-editor')).toBeVisible()
  await page.locator('.stage-wrap').click({ position: { x: 6, y: 6 } })
  await page.keyboard.press('F5')
  await page.locator('.slideshow, .presenter').first().waitFor()
  if (!(await page.locator('.presenter').count())) {
    await page.locator('.ss-controls button').last().click({ force: true })
    await page.getByText('Use Presenter View', { exact: true }).click()
  }
  await expect(page.locator('.pv-top')).toBeVisible()
}

test('presenter toolbar has full-size targets and its padded edges activate commands', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'presenter-controls' })
  try {
    const hasOverlay = await launched.app.evaluate(({ BrowserWindow, nativeTheme }) => {
      if (process.platform === 'darwin') return false
      const win = BrowserWindow.getAllWindows()[0]
      const original = win.setTitleBarOverlay.bind(win)
      win.setTitleBarOverlay = (options) => {
        Object.assign(globalThis, { presenterTestOverlay: options })
        original(options)
      }
      nativeTheme.themeSource = 'light'
      return true
    })
    const overlay = () =>
      launched.app.evaluate(
        () =>
          (
            globalThis as typeof globalThis & {
              presenterTestOverlay?: { color?: string; symbolColor?: string }
            }
          ).presenterTestOverlay,
      )
    await launched.page.locator('.quick-card').nth(2).click()
    const page = await waitForPageWithUrl(launched.app, '://slides/', 20000)
    await expect(page.locator('.ribbon')).toBeVisible()
    await enterPresenter(page)
    // The parent renderer's drag region is native window state, beyond elementFromPoint.
    await expect(launched.page.locator('.tab-bar')).toBeHidden()
    if (hasOverlay) {
      await expect.poll(overlay).toMatchObject({ color: '#000000', symbolColor: '#ffffff' })
      // Changing the editor theme during a show must not recolor the caption buttons.
      await launched.app.evaluate(({ nativeTheme }) => {
        nativeTheme.themeSource = 'dark'
      })
      await expect.poll(overlay).toMatchObject({ color: '#000000', symbolColor: '#ffffff' })
    }
    const controls = page.locator('.pv-top-btn')
    await expect(controls).toHaveCount(3)
    for (const button of await controls.all()) {
      const geometry = await button.evaluate((node) => {
        const box = node.getBoundingClientRect()
        const points = [
          [3, 3],
          [box.width - 3, 3],
          [3, box.height - 3],
          [box.width - 3, box.height - 3],
          [box.width / 2, box.height / 2],
        ]
        return {
          width: box.width,
          height: box.height,
          appRegion: getComputedStyle(node).getPropertyValue('-webkit-app-region'),
          hits: points.map(
            ([x, y]) => document.elementFromPoint(box.x + x, box.y + y)?.closest('button') === node,
          ),
        }
      })
      // Windows fractional display scaling can round the CSS box by hundredths of a pixel.
      expect(Math.round(geometry.width)).toBeGreaterThanOrEqual(88)
      expect(Math.round(geometry.height)).toBeGreaterThanOrEqual(48)
      expect(geometry.appRegion).toBe('no-drag')
      expect(geometry.hits.every(Boolean)).toBe(true)
    }
    await page
      .locator('.pv-top')
      .screenshot({ path: test.info().outputPath('presenter-toolbar.png') })
    // Click the blank padding at the lower-right edge, away from the icon and label.
    const show = controls.nth(2)
    const size = await show.boundingBox()
    await show.click({ position: { x: size!.width - 3, y: size!.height - 3 } })
    await expect(page.locator('.presenter')).toHaveCount(0)
    await expect(page.locator('.slideshow')).toBeVisible()
    await expect(launched.page.locator('.tab-bar')).toBeHidden()
    await page.keyboard.press('Escape')
    await expect(page.locator('.slideshow')).toHaveCount(0)
    await expect(launched.page.locator('.tab-bar')).toBeVisible()
    if (hasOverlay) {
      await expect.poll(overlay).toMatchObject({ color: '#2a2a2a', symbolColor: '#e4e4e4' })
      await launched.app.evaluate(({ nativeTheme }) => {
        nativeTheme.themeSource = 'light'
      })
      await expect.poll(overlay).toMatchObject({ color: '#ebebeb', symbolColor: '#454746' })
    }
    await enterPresenter(page)
    const end = page.locator('.pv-top-exit')
    await end.focus()
    await page.keyboard.press('Tab')
    await page.keyboard.press('Shift+Tab')
    await expect(end).toHaveCSS('outline-style', 'solid')
    await end.click({ position: { x: 3, y: 3 } })
    await expect(page.locator('.presenter')).toHaveCount(0)
    await expect(page.locator('.ribbon')).toBeVisible()
    await expect(launched.page.locator('.tab-bar')).toBeVisible()
    if (hasOverlay) {
      await expect.poll(overlay).toMatchObject({ color: '#ebebeb', symbolColor: '#454746' })
    }
  } finally {
    await closeAndSaveVideo(launched, 'presenter-controls')
  }
})
