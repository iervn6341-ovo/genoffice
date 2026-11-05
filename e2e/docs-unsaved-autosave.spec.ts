import { expect, test } from '@playwright/test'
import { mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('editing an unsaved blank Docs document does not create a file on background recovery', async () => {
  test.setTimeout(90_000)
  const root = mkdtempSync(join(tmpdir(), 'genoffice-docs-unsaved-'))
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { defaultSaveDir: root },
    videoDir: 'docs-unsaved-autosave',
  })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const docs = await waitForPageWithUrl(launched.app, '://docs/', 20_000)
    await expect(docs.locator('.doc-page').first()).toBeVisible({ timeout: 20_000 })
    await docs.locator('.doc-page').first().click()
    await docs.keyboard.type('Unsaved draft', { delay: 20 })
    await expect(docs.locator('.doc-page').first()).toContainText('Unsaved draft')
    await docs.waitForTimeout(32_000)
    expect(readdirSync(root)).toEqual([])
    await launched.page.locator('.tab-bar .tab-item.tab-home').click()
    await expect(launched.page.locator('.recent-row')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'docs-unsaved-autosave')
    rmSync(root, { recursive: true, force: true })
  }
})
