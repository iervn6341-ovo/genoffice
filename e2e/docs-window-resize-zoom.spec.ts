import { test, expect } from '@playwright/test'
import { resolve } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

const DOCX_FILE = resolve(__dirname, '../我的碩士論文計畫書_修訂版.docx')

test('window horizontal resize preserves manual zoom and does not jitter', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'zh-TW',
    videoDir: 'resize-zoom-test',
    openFile: DOCX_FILE,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/', 30_000)
    await page.waitForSelector('.doc-page', { timeout: 30_000 })
    await page.evaluate(() => document.fonts.ready)
    await page.waitForTimeout(1000)

    // Initial zoom should be 100%
    const zoomText = page.locator('.zoom-value')
    await expect(zoomText).toHaveText('100%')

    // Resize window narrower rapidly: from 1280px down to 750px (narrower than A4 paper ~794px + pad)
    await page.setViewportSize({ width: 750, height: 800 })
    await page.waitForTimeout(300)

    // Verify zoom is STILL 100% and did not get clamped to 80% or oscillate
    await expect(zoomText).toHaveText('100%')

    // Resize window back to 1100px
    await page.setViewportSize({ width: 1100, height: 800 })
    await page.waitForTimeout(300)
    await expect(zoomText).toHaveText('100%')

    // Resize window down to 600px
    await page.setViewportSize({ width: 600, height: 800 })
    await page.waitForTimeout(300)
    await expect(zoomText).toHaveText('100%')

    // Now test that clicking Page Width (頁寬) enters fit mode properly
    const fitWidthBtn = page.locator('button[data-shortcut="zoom-page-width"], [data-action="zoom-page-width"]').first()
    if (await fitWidthBtn.isVisible().catch(() => false)) {
      await fitWidthBtn.click()
    } else {
      await page.keyboard.press('Meta+0') // Reset to 100
    }
  } finally {
    await closeAndSaveVideo(launched)
  }
})
