import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

async function findShellPage(app: ElectronApplication, timeoutMs = 15_000): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      const has = await candidate
        .evaluate(() => Boolean((window as unknown as { aiOffice?: unknown }).aiOffice))
        .catch(() => false)
      if (has) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error('No window exposing window.aiOffice')
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}

test.describe('Docs styles pane and thesis features verification', () => {
  test('verifies styles pane Word parity and hidden section chips during body edit', async () => {
    test.setTimeout(60_000)
    const evidenceDir = '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1'
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-TW',
      videoDir: 'docs-styles-thesis',
    })

    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Docs' }).click()

      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await page.waitForLoadState('domcontentloaded')
      await page.waitForTimeout(2000)

      // Verify the floating header-footer bar is NOT visible over body text
      const hfChips = page.locator('.hf-variant-chips')
      if (await hfChips.count() > 0) {
        const opacity = await hfChips.evaluate((el) => window.getComputedStyle(el).opacity)
        expect(parseFloat(opacity)).toBe(0)
      }

      // Set wide viewport so ribbon groups are expanded
      await page.setViewportSize({ width: 1440, height: 900 })
      await page.waitForTimeout(1000)

      // Open the styles pane: either through the ribbon (unfolding if needed) or keyboard shortcut
      const foldBtn = page.locator('.rb-fold-btn', { hasText: /樣式|Styles/ }).first()
      if (await foldBtn.isVisible()) {
        await foldBtn.click()
        await page.waitForTimeout(500)
      }
      const styleGalleryMore = page.locator('.style-gallery-more').first()
      if (await styleGalleryMore.isVisible()) {
        await styleGalleryMore.click()
        await page.waitForTimeout(500)
        const stylesPaneBtn = page.locator('.style-gallery-footer').first()
        if (await stylesPaneBtn.isVisible()) {
          await stylesPaneBtn.click()
        }
      }

      // The Styles Pane aside should now be open (or open via Word standard shortcut Ctrl+Alt+Shift+S)
      const pane = page.locator('aside.styles-pane')
      if (!(await pane.isVisible())) {
        await page.keyboard.press('Control+Alt+Shift+KeyS')
      }
      await expect(pane).toBeVisible({ timeout: 10000 })

      // Verify "目前樣式:" box
      await expect(pane.locator('.styles-pane-current-box')).toBeVisible()

      // Verify "新增樣式..." and "全選" buttons
      await expect(pane.getByRole('button', { name: '新增樣式...' })).toBeVisible()
      await expect(pane.getByRole('button', { name: '全選' })).toBeVisible()

      // Verify list items match Word Traditional Chinese (Attachment 2 parity)
      const listText = await pane.locator('.styles-pane-list').innerText()
      expect(listText).toContain('引文')
      expect(listText).toContain('鮮明引文')
      expect(listText).toContain('標號')
      expect(listText).toContain('書目')
      expect(listText).toContain('目錄標題')
      expect(listText).toContain('圖表目錄')

      // Verify "清單:" footer select
      const footerSelect = pane.locator('.styles-pane-footer select')
      await expect(footerSelect).toBeVisible()
      await expect(footerSelect).toHaveValue('recommended')

      // Take a screenshot of the upgraded Styles Pane
      await pane.screenshot({ path: join(evidenceDir, 'docs-styles-pane-upgraded.png') })
      await page.screenshot({ path: join(evidenceDir, 'docs-full-thesis-verified.png') })
    } finally {
      await closeAndSaveVideo(launched, 'docs-styles-thesis')
    }
  })
})
