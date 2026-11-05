import { expect, test } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('all editor views follow the final native window size after a resize burst', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'window-resize',
    env: { GENOFFICE_NO_SPARE_VIEW: '' },
  })
  try {
    for (const name of ['Docs', 'Sheets', 'Slides', 'Markdown', 'HTML', 'PDF']) {
      await launched.page.locator('.tab-item.tab-home').click()
      await launched.page.locator('.quick-card', { hasText: `AI ${name}` }).click()
      const page = await waitForPageWithUrl(launched.app, `://${name.toLowerCase()}/`, 30_000)
      await expect(
        page.locator(name === 'PDF' ? '.pdf-page canvas' : '.ribbon').first(),
      ).toBeVisible({ timeout: 30_000 })
      if (name === 'Docs') {
        // Exercise the normal (non-test) prewarm policy: an active Docs tab
        // must not start a second, hidden editor during the idle timeout.
        await page.waitForTimeout(1700)
        expect(
          await launched.app.evaluate(
            ({ webContents }) =>
              webContents
                .getAllWebContents()
                .filter((contents) => contents.getURL().includes('://sheets/')).length,
          ),
        ).toBe(0)
      }
      const expected = await launched.app.evaluate(async ({ BrowserWindow }) => {
        const window = BrowserWindow.getAllWindows()[0]!
        for (let i = 0; i < 16; i++) {
          window.setSize(820 + (i % 8) * 50, 720 + (i % 3) * 20)
          await new Promise((resolve) => setTimeout(resolve, 12))
        }
        window.setSize(1280, 800)
        const bounds = window.getContentBounds()
        return { width: bounds.width, height: bounds.height - 40 }
      })
      await expect
        .poll(() => page.evaluate(() => ({ width: innerWidth, height: innerHeight })))
        .toEqual(expected)
    }
  } finally {
    await closeAndSaveVideo(launched, 'window-resize')
  }
})
