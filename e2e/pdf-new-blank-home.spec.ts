import { expect, test } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('home opens a blank PDF and renders its page', async () => {
  test.setTimeout(60_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'pdf-new-blank-home' })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI PDF' }).click()
    const pdf = await waitForPageWithUrl(launched.app, '://pdf/', 20_000)
    await expect(pdf.locator('.pdf-page canvas').first()).toBeVisible({ timeout: 20_000 })
    await expect(pdf.locator('.pdf-page')).toHaveCount(1)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-new-blank-home')
  }
})
