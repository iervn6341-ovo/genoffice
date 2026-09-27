import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

/**
 * PowerPoint for Mac (observed): with a slide thumbnail selected, ⌘⇧D inserts a copy of
 * that slide right after it; ⌘Z removes the copy again.
 */

const slideNodeCounts = (s: Page) =>
  s.evaluate(async () =>
    ((await (window as any).slidesApi.getRenderSlides()) as any[]).map((x) => x.nodes.length),
  )

test('⌘⇧D duplicates the selected slide after it; ⌘Z undoes it', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-duplicate-shortcut',
  })
  try {
    await launched.page.locator('.quick-card').nth(2).click()
    const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
    await s.waitForSelector('.ribbon', { timeout: 15_000 })
    await setEditorLayoutWidth(launched.app, '://slides/', 2200)
    // slide 1 gets a text box so its copy is recognisable; slide 2 stays empty
    await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
    await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
    await s.waitForSelector('.slide-text-editor')
    await s.keyboard.type('first')
    await s.keyboard.press('Escape')
    await s.keyboard.press('Escape')
    await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
    await s.locator('.ribbon-body button', { hasText: 'New Slide' }).first().click()
    const thumbs = s.locator('.slide-list .thumb')
    await expect(thumbs).toHaveCount(2)
    const before = await slideNodeCounts(s)

    await thumbs.nth(0).click()
    await s.keyboard.press('ControlOrMeta+Shift+d')
    await expect(thumbs).toHaveCount(3)
    // the copy sits right after the original
    await expect.poll(() => slideNodeCounts(s)).toEqual([before[0], before[0], before[1]])

    await s.keyboard.press('ControlOrMeta+z')
    await expect(thumbs).toHaveCount(2)
    await expect.poll(() => slideNodeCounts(s)).toEqual(before)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
