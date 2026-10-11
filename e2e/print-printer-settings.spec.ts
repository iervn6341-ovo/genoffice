import { test, expect, type Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Docs and Sheets print sheets carry the Office/macOS printer block: printer,
 * copies, two-sided and collated (Slides shares the same component).
 */

async function pressPrintMenu(app: LaunchedApp['app']): Promise<void> {
  await app.evaluate(({ Menu, BrowserWindow }) => {
    const walk = (m: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
      for (const it of m?.items ?? []) {
        if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(String(it.accelerator ?? ''))) {
          return it
        }
        const r = walk(it.submenu)
        if (r) return r
      }
    }
    walk(Menu.getApplicationMenu())?.click(
      undefined,
      BrowserWindow.getAllWindows()[0],
      undefined as never,
    )
  })
}

async function expectPrinterBlock(p: Page): Promise<void> {
  const block = p.locator('.printer-settings')
  await expect(block).toBeVisible({ timeout: 15_000 })
  await expect(block.locator('select.ps-select').first()).toBeVisible()
  await expect(block.locator('input.ps-num')).toHaveCount(1)
  await expect(block.locator('.ps-check input[type="checkbox"]')).toHaveCount(1)
  // two-sided offers off / long edge / short edge
  await expect(block.locator('label.ps-field select option')).toHaveCount(3)
}

test('docs: print sheet has printer / copies / two-sided / collated', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'print-docs' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForSelector('.doc-page', { timeout: 30_000 })
    await pressPrintMenu(launched.app)
    await expectPrinterBlock(p)
  } finally {
    await closeAndSaveVideo(launched, 'print-docs')
  }
})

test('sheets: print sheet has printer / copies / two-sided / collated', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'print-sheets' })
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const s = await waitForPageWithUrl(launched.app, '://sheets/')
    await s.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await s.waitForTimeout(1_500)
    await pressPrintMenu(launched.app)
    await expectPrinterBlock(s)
    // Escape closes the sheet without printing
    await s.keyboard.press('Escape')
    await expect(s.locator('.printer-settings')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'print-sheets')
  }
})
