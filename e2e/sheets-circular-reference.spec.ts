import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Excel (observed on Mac) reports a circular reference in the status bar —
 * "Circular References: A1" — and drops the notice once the cycle is gone.
 */

const api = (p: Page, body: string) =>
  p.evaluate(
    `(async () => { const wb = window.__genofficeDebug.univerAPI.getActiveWorkbook(); const sh = wb.getActiveSheet(); ${body} })()`,
  )
const indicator = (p: Page) => p.locator('.status-circular')

test('status bar names a circular reference and clears it when fixed', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-circular' })
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const s = await waitForPageWithUrl(launched.app, '://sheets/')
    await s.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await s.waitForTimeout(1_500)
    await api(
      s,
      `await sh.getRange(0, 0).setValue(5); await sh.getRange(0, 1).setValue('=A1*2'); return 1`,
    )
    await s.waitForTimeout(800)
    await expect(indicator(s)).toHaveCount(0)

    // two-cell loop D4 ↔ E4: the top-left cell is named
    await api(
      s,
      `await sh.getRange(3, 3).setValue('=E4+1'); await sh.getRange(3, 4).setValue('=D4+1'); return 1`,
    )
    await expect(indicator(s)).toHaveText('Circular References: D4')

    // break it: the notice goes away
    await api(s, `await sh.getRange(3, 4).setValue(7); return 1`)
    await expect(indicator(s)).toHaveCount(0)

    // a self-reference
    await api(s, `await sh.getRange(2, 2).setValue('=C3'); return 1`)
    await expect(indicator(s)).toHaveText('Circular References: C3')
    await api(s, `await sh.getRange(2, 2).setValue(''); return 1`)
    await expect(indicator(s)).toHaveCount(0)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
