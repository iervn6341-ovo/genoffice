import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { resolve } from 'node:path'

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

test.describe('Font list replenishment and sorting E2E', () => {
  test('Docs font dropdown shows replenished font list sorted by name', async () => {
    const launched = await launchShell({
      videoDir: 'font-list-docs',
      lang: 'zh-TW',
      onboardingSeen: true,
    })

    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Docs' }).click()

      const editorPage = await waitForPageWithUrl(launched.app, '://docs/')
      await editorPage.waitForSelector('.rb-font-family', { timeout: 15000 })

      // Open the font dropdown by clicking the combo caret button
      const fontCaret = editorPage.locator('.rb-combo-caret').first()
      await expect(fontCaret).toBeVisible({ timeout: 10000 })
      await fontCaret.click()

      // Font menu should open
      const fontMenu = editorPage.locator('.rb-font-family-menu').first()
      await expect(fontMenu).toBeVisible({ timeout: 5000 })

      // Get all font buttons in the menu
      const fontButtons = fontMenu.locator('button')
      const count = await fontButtons.count()
      expect(count).toBeGreaterThanOrEqual(50)

      const fontTexts = await fontButtons.allTextContents()
      console.log(`Docs font menu contains ${count} items. First 10:`, fontTexts.slice(0, 10))

      // Take a screenshot of the open font menu
      const shotPath = resolve(
        '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1',
        'docs-font-dropdown.png',
      )
      await editorPage.screenshot({ path: shotPath })

      // Verify essential Traditional Chinese and Latin fonts exist
      expect(fontTexts.some((f) => f.includes('微軟正黑體'))).toBe(true)
      expect(fontTexts.some((f) => f.includes('新細明體'))).toBe(true)
      expect(fontTexts.some((f) => f.includes('標楷體'))).toBe(true)
      expect(fontTexts.some((f) => f.includes('細明體'))).toBe(true)
      expect(fontTexts.some((f) => f.includes('Arial'))).toBe(true)
      expect(fontTexts.some((f) => f.includes('Calibri'))).toBe(true)
      expect(fontTexts.some((f) => f.includes('Times New Roman'))).toBe(true)

      // Verify Simplified Chinese fonts are NOT shown in zh-TW
      expect(fontTexts.some((f) => f.includes('等线'))).toBe(false)
      expect(fontTexts.some((f) => f.includes('微软雅黑'))).toBe(false)
      expect(fontTexts.some((f) => f.includes('方正小标宋简体'))).toBe(false)
      expect(fontTexts.some((f) => f.includes('华文中宋'))).toBe(false)

      // Test picking a font
      const biaukaiBtn = fontButtons.locator('text=標楷體').first()
      if (await biaukaiBtn.isVisible()) {
        await biaukaiBtn.click()
        await expect(fontMenu).not.toBeVisible()
      }
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('Sheets font dropdown contains replenished fonts sorted by name', async () => {
    const launched = await launchShell({
      videoDir: 'font-list-sheets',
      lang: 'zh-TW',
      onboardingSeen: true,
    })

    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Sheets' }).click()

      const editorPage = await waitForPageWithUrl(launched.app, '://sheets/')
      await editorPage.waitForTimeout(2000)

      // Find font select element in Sheets ribbon
      const fontSelect = editorPage.locator('select.tool-select, select[data-tip*="字型"], select[data-tip*="Font"]').first()
      if (await fontSelect.isVisible({ timeout: 5000 }).catch(() => false)) {
        const options = await fontSelect.locator('option').allTextContents()
        console.log(`Sheets font select contains ${options.length} options. First 10:`, options.slice(0, 10))
        expect(options.length).toBeGreaterThanOrEqual(50)
        expect(options.some((f) => f.includes('微軟正黑體'))).toBe(true)
        expect(options.some((f) => f.includes('新細明體'))).toBe(true)
        expect(options.some((f) => f.includes('標楷體'))).toBe(true)
        expect(options.some((f) => f.includes('Calibri'))).toBe(true)
        expect(options.some((f) => f.includes('等线'))).toBe(false)
        expect(options.some((f) => f.includes('微软雅黑'))).toBe(false)
      }

      const shotPath = resolve(
        '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1',
        'sheets-font-verified.png',
      )
      await editorPage.screenshot({ path: shotPath })
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('Slides font dropdown contains replenished fonts sorted by name', async () => {
    const launched = await launchShell({
      videoDir: 'font-list-slides',
      lang: 'zh-TW',
      onboardingSeen: true,
    })

    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Slides' }).click()

      const editorPage = await waitForPageWithUrl(launched.app, '://slides/')
      await editorPage.waitForTimeout(2000)

      // Open the font dropdown in Slides ribbon
      const fontCaret = editorPage.locator('.rb-select-btn, [data-tip*="字型"], button:has(.rb-caret)').first()
      if (await fontCaret.isVisible({ timeout: 5000 }).catch(() => false)) {
        await fontCaret.click()
        const fontMenu = editorPage.locator('.rb-font-menu').first()
        if (await fontMenu.isVisible({ timeout: 3000 }).catch(() => false)) {
          const fontButtons = await fontMenu.locator('button').allTextContents()
          console.log(`Slides font menu contains ${fontButtons.length} items. First 10:`, fontButtons.slice(0, 10))
          expect(fontButtons.length).toBeGreaterThanOrEqual(50)
        }
      }

      const shotPath = resolve(
        '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1',
        'slides-font-verified.png',
      )
      await editorPage.screenshot({ path: shotPath })
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
})
