import { test, expect } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/** Docs print sheet grouped like Word's print pane: Printer / Settings / Page Setup / Options. */
test('docs: print sheet groups, page setup and pages per sheet', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'docs-print' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForSelector('.doc-page', { timeout: 30_000 })
    await p.locator('.ProseMirror').click()
    // four pages: a hard page break after each line
    for (const n of [1, 2, 3, 4]) {
      await p.keyboard.type(`Page ${n}`)
      if (n < 4) await p.keyboard.press('ControlOrMeta+Enter')
    }
    // macOS prints through the system panel; this spec drives the in-app dialog (Windows/Linux entry)
    await p.evaluate(() => localStorage.setItem('genoffice:docs-force-in-app-print', '1'))
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
    await p.waitForSelector('.print-dialog', { timeout: 15_000 })
    await expect(p.locator('.print-section-head')).toHaveText([
      'Printer',
      'Settings',
      'Page Setup',
      'Options',
    ])
    const count = p.locator('.print-page-count')
    await expect(count).toContainText('4', { timeout: 20_000 })

    // odd pages only → 2; reverse order keeps the count
    await p.locator('[data-testid="print-parity"]').selectOption('odd')
    await expect(count).toContainText('2')
    await p.locator('[data-testid="print-parity"]').selectOption('even')
    await expect(count).toContainText('2')
    await p.locator('[data-testid="print-parity"]').selectOption('all')
    await expect(count).toContainText('4')

    // custom range
    await p.locator('[data-testid="print-range"]').selectOption('custom')
    await p.locator('.print-range-input').fill('2-3')
    await expect(count).toContainText('2')
    await p.locator('[data-testid="print-range"]').selectOption('all')

    // page setup edits the document: landscape swaps the sheet shape in the preview
    await p.locator('.print-section-head', { hasText: 'Page Setup' }).click()
    const box = () => p.locator('.print-preview-zoom .pv-sheet').first().boundingBox()
    const before = await box()
    expect(before!.height).toBeGreaterThan(before!.width)
    await p.locator('[data-testid="print-orientation"]').selectOption('landscape')
    await expect
      .poll(async () => (await box())!.width > (await box())!.height, { timeout: 20_000 })
      .toBe(true)
    await p.locator('[data-testid="print-orientation"]').selectOption('portrait')
    await expect
      .poll(async () => (await box())!.height > (await box())!.width, { timeout: 20_000 })
      .toBe(true)

    // paper size and margins presets
    await p.locator('[data-testid="print-paper"]').selectOption('letter')
    await expect(p.locator('[data-testid="print-paper"]')).toHaveValue('letter')
    await p.locator('[data-testid="print-margins"]').selectOption('narrow')
    await expect(p.locator('[data-testid="print-margins"]')).toHaveValue('narrow')

    // 4 pages per sheet → one sheet holding four pages
    await p.locator('[data-testid="print-per-sheet"]').selectOption('4')
    await expect(count).toContainText('1')
    await expect(p.locator('.print-preview-zoom .pv-sheet .pv-cell')).toHaveCount(4)
    // scale to paper size: the sheet takes that paper's proportions
    await p.locator('[data-testid="print-per-sheet"]').selectOption('1')
    await p.locator('[data-testid="print-scale-paper"]').selectOption('a3')
    await expect
      .poll(async () => {
        const b = (await box())!
        return Math.round((b.height / b.width) * 100) / 100
      })
      .toBeCloseTo(1.41, 1)

    await p.locator('.print-section-head', { hasText: 'Options' }).click()
    await expect(p.locator('.print-section-body .ps-check', { hasText: 'Reverse' })).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-print')
  }
})
