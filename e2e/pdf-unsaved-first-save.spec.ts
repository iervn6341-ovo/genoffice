import { expect, test } from '@playwright/test'
import { existsSync, mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('blank PDF creates a file and recent entry only after confirmed Save', async () => {
  test.setTimeout(90_000)
  const root = mkdtempSync(join(tmpdir(), 'genoffice-pdf-first-save-'))
  const target = join(root, 'My PDF.pdf')
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { defaultSaveDir: root },
    videoDir: 'pdf-unsaved-first-save',
  })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI PDF' }).click()
    const pdf = await waitForPageWithUrl(launched.app, '://pdf/', 20_000)
    await expect(pdf.locator('.pdf-page canvas').first()).toBeVisible({ timeout: 20_000 })
    expect(readdirSync(root)).toEqual([])

    await launched.app.evaluate(({ dialog }, chosen) => {
      ;(globalThis as any).__pdfFirstSaveDefaults = [] as string[]
      let attempts = 0
      dialog.showSaveDialog = (async (...args: unknown[]) => {
        const options = args.at(-1) as { defaultPath?: string }
        ;(globalThis as any).__pdfFirstSaveDefaults.push(options.defaultPath ?? '')
        attempts++
        return attempts === 1
          ? { canceled: true, filePath: undefined }
          : { canceled: false, filePath: chosen }
      }) as never
    }, target)

    await pdf.evaluate(() => (window as any).pdfApi.requestInitialSave())
    expect(existsSync(target)).toBe(false)
    await pdf.evaluate(() => {
      void (window as any).pdfApi.requestInitialSave()
    })
    await expect.poll(() => existsSync(target), { timeout: 20_000 }).toBe(true)
    const defaults = await launched.app.evaluate(() => (globalThis as any).__pdfFirstSaveDefaults)
    expect(defaults).toEqual([join(root, 'Untitled PDF.pdf'), join(root, 'Untitled PDF.pdf')])
    await launched.page.locator('.tab-bar .tab-item.tab-home').click()
    await expect(launched.page.locator('.recent-row', { hasText: 'My PDF.pdf' })).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'pdf-unsaved-first-save')
    rmSync(root, { recursive: true, force: true })
  }
})
