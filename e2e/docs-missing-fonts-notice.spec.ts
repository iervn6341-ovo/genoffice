import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { join, resolve } from 'node:path'
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

test.describe('Docs missing fonts auto popup and notification banner', () => {
  test('verifies missing fonts auto popup modal, table mapping, banner, and remember choice', async () => {
    test.setTimeout(60_000)
    const evidenceDir = '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1'
    const thesisDocxPath = resolve(process.cwd(), './我的碩士論文計畫書_修訂版.docx')

    console.log('[Test] Launching shell...')
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-TW',
      videoDir: 'docs-missing-fonts',
    })

    try {
      console.log('[Test] Finding shell page...')
      const shellPage = await findShellPage(launched.app)
      console.log('[Test] Clicking AI Docs card...')
      await shellPage.locator('.quick-card', { hasText: 'AI Docs' }).click()

      console.log('[Test] Waiting for Docs window...')
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await page.waitForLoadState('domcontentloaded')
      await page.waitForTimeout(2000)
      await page.setViewportSize({ width: 1440, height: 900 })

      // Clear any previous suppressed flag
      await page.evaluate((path) => {
        localStorage.removeItem(`aidocs.suppressFontModal.${path}`)
      }, thesisDocxPath)

      console.log('[Test] Opening thesis document with uninstalled fonts...')
      await page.evaluate(async (docPath) => {
        const aidocs = (window as any).__aidocs
        if (aidocs?.openPath) {
          await aidocs.openPath(docPath)
        }
      }, thesisDocxPath)

      // 1. Verify that MissingFontsDialog automatically pops up upon document open!
      console.log('[Test] Waiting for MissingFontsDialog auto popup...')
      const modal = page.locator('.missing-fonts-modal')
      await expect(modal).toBeVisible({ timeout: 15_000 })

      // Check title and table contents
      await expect(modal.locator('h2')).toContainText('文件字型相容性提示')
      const rows = modal.locator('.missing-fonts-table tbody tr')
      const rowCount = await rows.count()
      console.log('[Test] Missing font rows count:', rowCount)
      expect(rowCount).toBeGreaterThan(0)

      // Verify 新細明體 or 細明體 mapped to GenOffice Songti TC
      const tableText = await modal.locator('.missing-fonts-table').textContent()
      console.log('[Test] Table content snippet:', tableText?.slice(0, 100))
      expect(tableText).toContain('新細明體')
      expect(tableText).toContain('GenOffice Songti TC')

      // Capture screenshot of auto popup dialog
      await page.screenshot({ path: join(evidenceDir, 'docs-missing-fonts-dialog-verified.png') })
      console.log('[Test] Saved missing fonts dialog screenshot.')

      // 2. Click "我知道了" button to dismiss modal
      const confirmBtn = modal.locator('.missing-fonts-confirm-btn')
      await confirmBtn.click()
      await page.waitForTimeout(400)
      await expect(modal).toBeHidden()

      // 3. Verify top notification banner (MissingFontsBanner) below Ribbon is visible
      const banner = page.locator('.missing-fonts-banner')
      await expect(banner).toBeVisible()
      await expect(banner).toContainText('新細明體')

      // Capture screenshot of banner below ribbon
      await page.screenshot({ path: join(evidenceDir, 'docs-missing-fonts-banner-verified.png') })
      console.log('[Test] Saved missing fonts banner screenshot.')

      // 4. Verify bottom status bar message is clickable
      const statusMsg = page.locator('.status-font-clickable')
      await expect(statusMsg).toBeVisible()
      await expect(statusMsg).toContainText('新細明體')

      // Click status bar message and verify modal re-opens
      console.log('[Test] Clicking status bar font message to re-open dialog...')
      await statusMsg.click()
      await page.waitForTimeout(400)
      await expect(modal).toBeVisible()

      // 5. Test "Don't show again" checkbox
      const dontShowCheckbox = modal.locator('.missing-fonts-checkbox-label input[type="checkbox"]')
      await dontShowCheckbox.check()
      await expect(dontShowCheckbox).toBeChecked()

      // Confirm and close
      await modal.locator('.missing-fonts-confirm-btn').click()
      await page.waitForTimeout(400)
      await expect(modal).toBeHidden()

      // Dismiss banner
      await banner.locator('.missing-fonts-banner-dismiss-btn').click()
      await page.waitForTimeout(300)
      await expect(banner).toBeHidden()

      console.log('[Test] All missing fonts prompt assertions verified successfully!')
    } finally {
      await closeAndSaveVideo(launched, 'docs-missing-fonts')
    }
  })
})
