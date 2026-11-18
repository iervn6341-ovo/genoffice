import { expect, test } from '@playwright/test'
import { mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell } from './helpers'

test('opening blank templates does not create user documents', async () => {
  test.setTimeout(120_000)
  const root = mkdtempSync(join(tmpdir(), 'genoffice-blank-no-file-'))
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { defaultSaveDir: root },
    videoDir: 'home-blank-no-file',
  })
  try {
    let tabCount = 1
    for (const kind of ['AI Docs', 'AI Sheets', 'AI Slides', 'AI Markdown', 'AI HTML', 'AI PDF']) {
      await launched.page.locator('.quick-card', { hasText: kind }).click()
      tabCount++
      await expect(launched.page.locator('.tab-bar .tab-item')).toHaveCount(tabCount, {
        timeout: 20_000,
      })
      await launched.page.locator('.tab-bar .tab-item.tab-home').click()
      expect(readdirSync(root)).toEqual([])
      await expect(launched.page.locator('.recent-row')).toHaveCount(0)
    }
  } finally {
    await closeAndSaveVideo(launched, 'home-blank-no-file')
    rmSync(root, { recursive: true, force: true })
  }
})
