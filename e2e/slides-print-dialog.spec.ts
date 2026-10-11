import { test, expect } from '@playwright/test'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/** Print sheet grouped like Office's print pane: Printer / Settings / Page Setup / Options. */
test('slides: print dialog options (range, layout, paper, output)', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'print-dialog',
    openFile: join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
  })
  try {
    const s = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await s.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    await launched.app.evaluate(({ Menu, BrowserWindow }) => {
      const walk = (m: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
        for (const it of m?.items ?? []) {
          const acc = String(it.accelerator ?? '')
          if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(acc)) return it
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
    await s.waitForSelector('.print-dialog', { timeout: 15_000 })
    const count = s.locator('.print-page-count')
    await expect(count).toContainText('5')
    const preview = (needle: string) =>
      s.evaluate(
        (n) =>
          (
            document.querySelector('iframe[title="print-preview"]') as HTMLIFrameElement | null
          )?.srcdoc.includes(n),
        needle,
      )
    const open = async (title: string) => {
      const head = s.locator('.print-section-head', { hasText: title })
      if ((await head.getAttribute('aria-expanded')) !== 'true') await head.click()
    }

    // the Office groups: Printer / Settings are open, Page Setup / Options start folded
    await expect(s.locator('.print-section-head')).toHaveText([
      'Printer',
      'Settings',
      'Page Setup',
      'Options',
    ])
    await expect(s.locator('.print-dialog .printer-settings input.ps-num')).toHaveCount(1)

    // range: custom "2-3" → two pages; a reversed range is invalid; selection / current work too
    await s.locator('[data-testid="print-range"]').selectOption('custom')
    const range = s.locator('.print-dialog .print-range-input')
    await range.fill('2-3')
    await expect(count).toContainText('2')
    await range.fill('4-2')
    await expect(count).toHaveText('')
    await range.fill('2-3')
    await s.locator('[data-testid="print-range"]').selectOption('current')
    await expect(count).toContainText('1')
    await s.locator('[data-testid="print-range"]').selectOption('all')
    await expect(count).toContainText('5')

    // pages per sheet: 4 slides on one sheet; direction buttons wake up
    await open('Page Setup')
    await s.locator('[data-testid="print-per-sheet"]').selectOption('4')
    await expect(count).toContainText('2')
    await expect(s.locator('.print-dir').nth(2)).toBeEnabled()
    await s.locator('.print-dir').nth(2).click()
    await expect.poll(() => preview('class="page nup"')).toBe(true)
    await s.locator('[data-testid="print-per-sheet"]').selectOption('1')

    // every layout of the Print Layout list: handouts of 2 per page, 9 vertical, outline, notes
    const layout = s.locator('[data-testid="print-layout"]')
    await layout.selectOption('handout2')
    await expect(count).toContainText('3')
    await layout.selectOption('handout9v')
    await expect(count).toContainText('1')
    await expect.poll(() => preview('class="page handout h9 hg"')).toBe(true)
    await layout.selectOption('handout4h')
    await expect(count).toContainText('2')
    await layout.selectOption('outline')
    await expect.poll(() => preview('class="page flow outline"')).toBe(true)
    await layout.selectOption('notes')
    await expect(count).toContainText('5')
    await layout.selectOption('handout2')

    // paper size: Letter becomes the page size
    await s.locator('[data-testid="print-paper"]').selectOption('letter')
    await expect.poll(() => preview('size: 8.5in 11in')).toBe(true)

    // color: pure black & white filters the preview images
    await s.locator('[data-testid="print-color"]').selectOption('bw')
    await expect.poll(() => preview('contrast(1000%)')).toBe(true)

    // options: frame, scale to fit paper (full page on a real paper), hidden slides, header/footer link
    await open('Options')
    await layout.selectOption('full')
    await s.locator('[data-testid="print-paper"]').selectOption('a4')
    const scaleFit = s.locator('.ps-check', { hasText: 'Scale to Fit Paper' }).locator('input')
    await expect(scaleFit).toBeEnabled()
    await scaleFit.check()
    await expect.poll(() => preview('class="page onpaper fit"')).toBe(true)
    await expect(s.locator('.ps-link')).toHaveText('Header & Footer…')
  } finally {
    await closeAndSaveVideo(launched, 'print-dialog')
  }
})
