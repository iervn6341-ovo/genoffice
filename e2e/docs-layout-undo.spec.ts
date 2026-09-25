import { test, expect } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

/**
 * C7 — page layout shares Word's undo stack: type a word, switch orientation
 * to Landscape, ⌘Z restores Portrait and keeps the word, ⌘⇧Z re-applies
 * Landscape, a second ⌘Z/⌘Z removes the word.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const pageIsLandscape = (p: import('@playwright/test').Page) =>
  p.evaluate(() => {
    const page = document.querySelector('.doc-page') as HTMLElement | null
    if (!page) return false
    const r = page.getBoundingClientRect()
    return r.width > r.height
  })

test('orientation change is undone before the typed word', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-layout-undo' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(() => Boolean((window as unknown as DocsWindow).__aidocs?.editor))
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('word')
    expect(await pageIsLandscape(p)).toBe(false)

    await p.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await p.locator('.ribbon-body button', { hasText: 'Orientation' }).first().click()
    await p.getByText('Landscape', { exact: true }).first().click()
    await expect.poll(() => pageIsLandscape(p)).toBe(true)

    await p.locator('.doc-page').first().click()
    await p.keyboard.press('Meta+z')
    await expect.poll(() => pageIsLandscape(p)).toBe(false)
    await expect(p.locator('.doc-page').first()).toContainText('word')

    await p.keyboard.press('Meta+Shift+z')
    await expect.poll(() => pageIsLandscape(p)).toBe(true)

    await p.keyboard.press('Meta+z')
    await p.keyboard.press('Meta+z')
    await expect(p.locator('.doc-page').first()).not.toContainText('word')
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
