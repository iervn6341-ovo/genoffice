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

test.describe('Docs and Slides element dashed/solid frame parity', () => {
  test('verifies text editing displays dashed outline and container selection displays solid outline', async () => {
    test.setTimeout(60_000)
    const evidenceDir = '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1'
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-TW',
      videoDir: 'docs-dashed-frame',
    })

    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Docs' }).click()

      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await page.waitForLoadState('domcontentloaded')
      await page.waitForTimeout(2000)
      await page.setViewportSize({ width: 1440, height: 900 })

      // Go to Insert tab and insert a Text Box
      const insertTab = page.locator('.ribbon-tabs button', { hasText: '插入' })
      await insertTab.click()
      await page.waitForTimeout(500)

      const textBoxBtn = page.locator('button', { hasText: '文字方塊' }).first()
      await textBoxBtn.click()
      await page.waitForTimeout(1000)

      // 1. Text editing state: clicking into the text box must show a DASHED outline (Word/PowerPoint parity)
      const textbox = page.locator('.doc-textbox').first()
      await expect(textbox).toBeVisible({ timeout: 5000 })
      await textbox.click()
      await page.waitForTimeout(500)

      const outlineStyleEditing = await textbox.evaluate(
        (el) => window.getComputedStyle(el).outlineStyle,
      )
      expect(outlineStyleEditing).toBe('dashed')
      await page.screenshot({ path: join(evidenceDir, 'docs-element-dashed-verified.png') })

      // 2. Object selected state: pressing Escape exits text editing into container selection mode
      // which must show a SOLID outline (Word/PowerPoint parity)
      await page.keyboard.press('Escape')
      await page.waitForTimeout(500)

      const outlineStyleSelected = await textbox.evaluate(
        (el) => window.getComputedStyle(el).outlineStyle,
      )
      expect(outlineStyleSelected).toBe('solid')
      await page.screenshot({ path: join(evidenceDir, 'docs-element-solid-verified.png') })
    } finally {
      await closeAndSaveVideo(launched, 'docs-dashed-frame')
    }
  })
})
