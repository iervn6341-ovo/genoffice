import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Reopened (streamed) workbook: the Number Format box must name the selected cell's
 * format. Streamed cells land without the value/numfmt mutations the ribbon echo
 * listens for, so the box used to stay "General" on a currency cell until the
 * selection moved.
 */

const api = (p: Page, body: string) =>
  p.evaluate(
    `(async () => { const sh = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${body} })()`,
  )
const numberFormatBox = (p: Page) =>
  p.evaluate(() => {
    const el = [...document.querySelectorAll('button, [role=combobox]')].find((e) =>
      /^(General|Currency|Number|Accounting|Percentage|Short Date|Long Date|Time|Fraction|Scientific|Text)\b/.test(
        (e as HTMLElement).innerText.trim(),
      ),
    ) as HTMLElement | undefined
    return el?.innerText.trim() ?? null
  })

test('Number Format box names a reopened currency cell without moving the selection', async () => {
  const out = join(mkdtempSync(join(tmpdir(), 'sheets-numfmt-echo-')), 'money.xlsx')
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-numfmt-echo' })
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const s = await waitForPageWithUrl(launched.app, '://sheets/')
    await s.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await s.waitForTimeout(1_500)
    await launched.app.evaluate(({ dialog }, f) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: f })) as never
    }, out)
    await api(s, `await sh.getRange(4, 1).setValue(1200); sh.getRange(4, 1).activate(); return 1`)
    await s.keyboard.press('Control+Shift+Digit4') // Excel: Currency
    await expect.poll(() => numberFormatBox(s)).toBe('Currency')
    await s.locator('button[aria-label^="Save As"]').first().click()
    await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
  } finally {
    launched.app.process().kill('SIGKILL')
  }

  const re = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-numfmt-echo',
    openFile: out,
  })
  try {
    const q = await waitForPageWithUrl(re.app, '://sheets/')
    await q.waitForFunction(
      () => Boolean((window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()),
      null,
      { timeout: 30_000 },
    )
    // select once, right away, and never move again
    await api(q, `sh.getRange(4, 1).activate(); return 1`)
    await expect.poll(() => numberFormatBox(q), { timeout: 8_000 }).toBe('Currency')
  } finally {
    re.app.process().kill('SIGKILL')
  }
})
