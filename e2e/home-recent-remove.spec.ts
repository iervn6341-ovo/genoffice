import { expect, test } from '@playwright/test'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell } from './helpers'

test('remove from recent list leaves the document on disk', async () => {
  const root = mkdtempSync(join(tmpdir(), 'genoffice-recent-remove-'))
  const filePath = join(root, 'keep.md')
  writeFileSync(filePath, '# Keep this file')
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { defaultSaveDir: root },
    videoDir: 'home-recent-remove',
  })
  try {
    await launched.page.evaluate((path) => window.aiOffice.openPath(path), filePath)
    await launched.page.locator('.tab-bar .tab-item.tab-home').click()
    const row = launched.page.locator('.recent-row', { hasText: 'keep.md' })
    await expect(row).toBeVisible()
    await row.locator('.remove-recent-btn').click()
    await expect(row).toHaveCount(0)
    expect(existsSync(filePath)).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'home-recent-remove')
    rmSync(root, { recursive: true, force: true })
  }
})
