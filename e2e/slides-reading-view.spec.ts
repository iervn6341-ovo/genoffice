import { test, expect } from '@playwright/test'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Reading View is PowerPoint's windowed slide show: the real show (animations, links, ink)
 * a bare window (no ribbon / status bar) like PowerPoint's, switched from the status bar's view buttons; Esc returns
 * to Normal.
 */
test('slides: status-bar Reading View plays the show inside the window', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'reading-view',
    openFile: join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
  })
  try {
    const s = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await s.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    const views = s.locator('.status-views .status-play-btn')
    await expect(views).toHaveCount(4)
    await expect(views.nth(0)).toHaveAttribute('aria-pressed', 'true')

    await views.nth(2).click()
    const show = s.locator('.slideshow.ss-windowed')
    await expect(show).toBeVisible()
    // a bare show window like PowerPoint's: no ribbon, no status bar, no side panes
    await expect(s.locator('.ribbon')).toBeHidden()
    await expect(s.locator('.status-bar')).toBeHidden()
    await expect(s.locator('.slide-list')).toHaveCount(0)
    const box = await show.boundingBox()
    const win = await s.evaluate(() => ({ width: innerWidth, height: innerHeight }))
    expect(box!.width).toBeGreaterThan(win.width * 0.95)
    expect(box!.height).toBeGreaterThan(win.height * 0.95)
    // the control bar is docked flush in the bottom-left corner
    const bar = await s.locator('.ss-controls').boundingBox()
    expect(bar!.x).toBeLessThan(2)
    expect(bar!.y + bar!.height).toBeGreaterThan(win.height - 2)
    await expect(s.locator('.ss-counter')).toContainText('1 /')
    await s.keyboard.press('ArrowRight')
    await expect(s.locator('.ss-counter')).toContainText('2 /')

    await s.keyboard.press('Escape')
    await expect(show).toHaveCount(0)
    await expect(s.locator('.stage-wrap')).toHaveCount(1)
    await expect(s.locator('.ribbon')).toBeVisible()
    await expect(s.locator('.status-bar')).toBeVisible()
    await expect(views.nth(0)).toHaveAttribute('aria-pressed', 'true')
  } finally {
    await closeAndSaveVideo(launched, 'reading-view')
  }
})

/**
 * Slide Sorter (status-bar button 2): the page number sits under each thumbnail, and the
 * status-bar zoom slider sizes the thumbnails like PowerPoint's.
 */
test('slides: slide sorter shows numbers under the pages and follows the zoom slider', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sorter-view',
    openFile: join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
  })
  try {
    const s = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await s.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    await s.locator('.status-views .status-play-btn').nth(1).click()
    const first = s.locator('.sorter-item').first()
    await expect(first).toBeVisible()
    // number is below the frame, not on top of it
    const frame = await first.locator('.sorter-frame').boundingBox()
    const num = await first.locator('.sorter-num').boundingBox()
    expect(num!.y).toBeGreaterThanOrEqual(frame!.y + frame!.height - 1)
    // the zoom slider resizes the thumbnails
    const before = (await first.locator('.sorter-frame').boundingBox())!.width
    await s.locator('.zoom-slider').fill('50')
    await expect
      .poll(async () => (await first.locator('.sorter-frame').boundingBox())!.width)
      .toBeLessThan(before)
  } finally {
    await closeAndSaveVideo(launched, 'sorter-view')
  }
})
