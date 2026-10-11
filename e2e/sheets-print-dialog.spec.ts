import { test, expect } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/** Sheets print sheet grouped like Excel's print pane: Printer / Settings / Page Setup / Options. */
test('sheets: print sheet scope, page setup and scaling', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'sheets-print' })
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const s = await waitForPageWithUrl(launched.app, '://sheets/')
    await s.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await s.waitForTimeout(1_500)
    await launched.app.evaluate(({ Menu, BrowserWindow }) => {
      const walk = (m: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
        for (const it of m?.items ?? []) {
          if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(String(it.accelerator ?? '')))
            return it
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
    await s.waitForSelector('.print-sheet-dialog', { timeout: 15_000 })
    await expect(s.locator('.print-section-head')).toHaveText([
      'Printer',
      'Settings',
      'Page Setup',
      'Options',
    ])

    // scope: active sheet / entire workbook / selection; pages from-to only for one sheet
    const scope = s.locator('[data-testid="print-scope"]')
    await expect(scope.locator('option')).toHaveText([
      'Print Active Sheets',
      'Print Entire Workbook',
      'Print Selection',
    ])
    await scope.selectOption('workbook')
    await expect(s.locator('[data-testid="print-page-from"]')).toBeDisabled()
    await scope.selectOption('active')
    await expect(s.locator('[data-testid="print-page-from"]')).toBeEnabled()
    await s.locator('[data-testid="print-page-from"]').fill('2')
    await s.locator('[data-testid="print-page-to"]').fill('3')

    // page setup edits the sheet's layout and the pane echoes it back
    await s.locator('.print-section-head', { hasText: 'Page Setup' }).click()
    const orientation = s.locator('[data-testid="print-orientation"]')
    await orientation.selectOption('landscape')
    await expect(orientation).toHaveValue('landscape')
    await expect.poll(async () => orientation.inputValue()).toBe('landscape')
    await s.locator('[data-testid="print-paper"]').selectOption('1')
    await expect.poll(async () => s.locator('[data-testid="print-paper"]').inputValue()).toBe('1')
    await s.locator('[data-testid="print-margins"]').selectOption('narrow')
    await expect
      .poll(async () => s.locator('[data-testid="print-margins"]').inputValue())
      .toBe('narrow')

    // scaling: fit sheet on one page, then a custom percentage
    const scaling = s.locator('[data-testid="print-scaling"]')
    await scaling.selectOption('fitSheet')
    await expect.poll(async () => scaling.inputValue()).toBe('fitSheet')
    await scaling.selectOption('fitCols')
    await expect.poll(async () => scaling.inputValue()).toBe('fitCols')
    await scaling.selectOption('custom')
    await expect.poll(async () => scaling.inputValue()).toBe('custom')
    await s.locator('[data-testid="print-scale"]').fill('80')
    await s.locator('[data-testid="print-scale"]').press('Enter')
    await expect.poll(async () => s.locator('[data-testid="print-scale"]').inputValue()).toBe('80')
    await scaling.selectOption('none')
    await expect.poll(async () => scaling.inputValue()).toBe('none')

    // options
    await s.locator('.print-section-head', { hasText: 'Options' }).click()
    const grid = s
      .locator('.print-section-body .ps-check', { hasText: 'Gridlines' })
      .locator('input')
    await grid.check()
    await expect.poll(async () => grid.isChecked()).toBe(true)
    await expect(s.locator('.ps-link')).toHaveText('Header & Footer…')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-print')
  }
})
