import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

/**
 * Thumbnail pane multi-selection, PowerPoint-style: Shift-click selects a range from the
 * last plain-clicked slide, ⌘-click toggles one slide, Delete removes the whole selection
 * as one undo step, and ⌘Z puts every slide back in its original place.
 * (Playwright's click `modifiers` don't reach this page — hold the key instead.)
 */

const slideTexts = (s: Page) =>
  s.evaluate(async () =>
    ((await (window as any).slidesApi.getRenderSlides()) as any[]).map((x) =>
      x.nodes
        .flatMap((n: any) => (n.text?.lines ?? []).flatMap((l: any) => l.runs))
        .map((r: any) => r.text)
        .join(''),
    ),
  )

test('Shift/⌘-click select thumbnails; Delete removes them in one undo step', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-thumb-multiselect' })
  try {
    await launched.page.locator('.quick-card').nth(2).click()
    const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
    await s.waitForSelector('.ribbon', { timeout: 15_000 })
    await setEditorLayoutWidth(launched.app, '://slides/', 2200)
    const thumbs = s.locator('.slide-list .thumb')
    for (let i = 1; i <= 4; i++) {
      if (i > 1) {
        await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
        await s.locator('.ribbon-body button', { hasText: 'New Slide' }).first().click()
        await expect(thumbs).toHaveCount(i)
      }
      await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
      await s.waitForSelector('.slide-text-editor')
      await s.keyboard.type(`S${i}`)
      await s.keyboard.press('Escape')
      await s.keyboard.press('Escape')
    }
    await expect.poll(() => slideTexts(s)).toEqual(['S1', 'S2', 'S3', 'S4'])

    await thumbs.nth(1).click()
    await s.keyboard.down('Shift')
    await thumbs.nth(2).click()
    await s.keyboard.up('Shift')
    await expect(s.locator('.slide-list .thumb.selected')).toHaveCount(2)
    await s.keyboard.down('Meta')
    await thumbs.nth(3).click() // add S4
    await thumbs.nth(2).click() // drop S3
    await s.keyboard.up('Meta')
    await expect(s.locator('.slide-list .thumb.selected')).toHaveCount(2)
    await expect(thumbs.nth(1)).toHaveClass(/selected/)
    await expect(thumbs.nth(3)).toHaveClass(/selected/)

    await s.keyboard.press('Delete')
    await expect.poll(() => slideTexts(s)).toEqual(['S1', 'S3'])
    await expect(s.locator('.slide-list .thumb.selected')).toHaveCount(0)

    // one ⌘Z brings both back, each in its original position
    await s.keyboard.press('ControlOrMeta+z')
    await expect.poll(() => slideTexts(s)).toEqual(['S1', 'S2', 'S3', 'S4'])

    // all four selected: a deck keeps at least one slide, so Delete does nothing
    await thumbs.nth(0).click()
    await s.keyboard.down('Shift')
    await thumbs.nth(3).click()
    await s.keyboard.up('Shift')
    await expect(s.locator('.slide-list .thumb.selected')).toHaveCount(4)
    await s.keyboard.press('Delete')
    await expect.poll(() => slideTexts(s)).toEqual(['S1', 'S2', 'S3', 'S4'])
    // a plain click collapses the selection; Delete then removes only that slide
    await thumbs.nth(2).click()
    await expect(s.locator('.slide-list .thumb.selected')).toHaveCount(0)
    await s.keyboard.press('Delete')
    await expect.poll(() => slideTexts(s)).toEqual(['S1', 'S2', 'S4'])
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
