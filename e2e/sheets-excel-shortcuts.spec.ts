import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Excel for Mac keyboard conventions on a new workbook:
 *  - ⇧⌘T (and ⌥=) AutoSum from the empty cell under a column of numbers;
 *  - Control+Shift+$ / % / ~ apply Currency / Percent / General;
 *  - the font box names the font unstyled cells really use (the workbook's Normal font).
 */

interface DebugWindow {
  __genofficeDebug: { univerAPI: any }
}

const sheet = (p: Page) =>
  p.evaluateHandle(() =>
    (window as unknown as DebugWindow).__genofficeDebug.univerAPI
      .getActiveWorkbook()
      .getActiveSheet(),
  )
const cell = (p: Page, a1: string) =>
  p.evaluate((ref) => {
    const r = (window as unknown as DebugWindow).__genofficeDebug.univerAPI
      .getActiveWorkbook()
      .getActiveSheet()
      .getRange(ref)
    return { formula: r.getFormula() as string, display: r.getDisplayValue() as string }
  }, a1)
const select = (p: Page, a1: string) =>
  p.evaluate(
    (ref) =>
      (window as unknown as DebugWindow).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(ref)
        .activate(),
    a1,
  )

test('AutoSum, number-format shortcuts and the default font echo match Excel', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-excel-shortcuts' })
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const p = await waitForPageWithUrl(launched.app, '://sheets/')
    await p.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await p.waitForTimeout(1_500)
    await sheet(p)
    const grid = await p.evaluate(() => {
      for (const c of document.querySelectorAll('canvas')) {
        const r = c.getBoundingClientRect()
        if (r.width > 500 && r.height > 300) return { x: r.x, y: r.y }
      }
      return null
    })
    await p.mouse.click(grid!.x + 46 + 43, grid!.y + 24 + 12)
    await select(p, 'A1')
    for (const v of ['3', '1.5', '0.25']) {
      await p.keyboard.type(v)
      await p.keyboard.press('Enter')
    }
    // the font box shows the workbook's Normal font, which is also what the file saves
    await expect(
      p.locator('input[aria-label="Font"], [aria-label="Font"] input').first(),
    ).toHaveValue('Calibri')

    await select(p, 'A4')
    await p.keyboard.press('Meta+Shift+t')
    await expect.poll(() => cell(p, 'A4')).toEqual({ formula: '=SUM(A1:A3)', display: '4.75' })

    await select(p, 'A1:A4')
    await p.keyboard.press('Control+Shift+4')
    await expect.poll(() => cell(p, 'A2').then((c) => c.display)).toBe('$1.50')
    await p.keyboard.press('Control+Shift+5')
    await expect.poll(() => cell(p, 'A3').then((c) => c.display)).toBe('25%')
    await p.keyboard.press('Control+Shift+`')
    await expect.poll(() => cell(p, 'A2').then((c) => c.display)).toBe('1.5')
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
