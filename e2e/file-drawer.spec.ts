import { test, expect } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * The File tab opens a left drawer (Office "backstage" as a slide-out menu) in Docs / Sheets / Slides:
 * New / Open / Info / Save / Save As / Print / Export, with the document still visible behind it.
 */
const APPS = [
  { app: 'docs', card: 0, label: 'Docs', exports: 3 },
  { app: 'sheets', card: 1, label: 'Sheets', exports: 2 },
  { app: 'slides', card: 2, label: 'Slides', exports: 2 },
] as const

for (const { app, card, label, exports } of APPS) {
  test(`${label}: File tab opens a left drawer`, async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      videoDir: 'file-drawer',
    })
    try {
      await launched.page.locator('.quick-card').nth(card).click()
      const s = await waitForPageWithUrl(launched.app, `://${app}/`, 30_000)
      await s.waitForSelector('.ribbon-tab-file', { timeout: 20_000 })
      await s.locator('.ribbon-tab-file').first().click()
      const drawer = s.locator('.bs-drawer')
      await expect(drawer).toBeVisible()
      // a drawer, not a page: the document area stays visible to its right
      const w = await drawer.evaluate((el) => el.getBoundingClientRect().width)
      const vw = await s.evaluate(() => innerWidth)
      expect(w).toBeLessThan(vw / 2)
      for (const name of ['New', 'Open', 'Info', 'Save', 'Save As', 'Print', 'Export']) {
        await expect(drawer.locator('.bs-nav', { hasText: new RegExp(`^${name}$`) })).toBeVisible()
      }
      // a pane replaces the menu inside the drawer; back (and Esc) returns to the menu
      await drawer.locator('.bs-nav', { hasText: /^Export$/ }).click()
      await expect(drawer.locator('.bs-export')).toHaveCount(exports)
      await s.keyboard.press('Escape')
      await expect(drawer.locator('.bs-nav').first()).toBeVisible()
      await drawer.locator('.bs-nav', { hasText: /^Open$/ }).click()
      await expect(drawer.locator('.bs-primary')).toBeVisible()
      await drawer.locator('.bs-subhead').click()
      // Esc on the menu closes the drawer; so does clicking the page behind it
      await s.keyboard.press('Escape')
      await expect(s.locator('.backstage')).toHaveCount(0)
      await s.locator('.ribbon-tab-file').first().click()
      await s.mouse.click(vw - 40, 400)
      await expect(s.locator('.backstage')).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'cleanup')
    }
  })
}
