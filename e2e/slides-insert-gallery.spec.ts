import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

async function openSlides(lang = 'en') {
  const launched = await launchShell({
    onboardingSeen: true,
    lang,
    videoDir: 'slides-insert-gallery',
  })
  await launched.page.locator('.quick-card').nth(2).click()
  const page = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await page.waitForSelector('.ribbon')
  return { launched, page }
}

const nodes = (page: Page) =>
  page.evaluate(
    async () => (await (window as any).slidesApi.getRenderSlides())[0].nodes.length as number,
  )

async function assertReservedSpace(page: Page) {
  await expect
    .poll(() =>
      page.evaluate(() => {
        const panel = document.querySelector('.ribbon-gallery-host')!.getBoundingClientRect()
        const stage = document.querySelector('.stage-col')!.getBoundingClientRect()
        return panel.left - stage.right
      }),
    )
    .toBeGreaterThanOrEqual(-1)
  const bounds = await page.locator('.stage-col').boundingBox()
  expect(bounds!.width).toBeGreaterThan(300)
  await expect
    .poll(() => page.evaluate(() => document.documentElement.scrollWidth - innerWidth))
    .toBeLessThanOrEqual(1)
}

test('chart gallery reserves canvas space, supports keyboard navigation and inserts undoable content', async () => {
  const { launched, page } = await openSlides()
  try {
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    const trigger = page.locator('[data-gallery-trigger="chart"]')
    const before = await nodes(page)
    const aiPreference = await page.evaluate(() => localStorage.getItem('ai-slides-show-ai'))
    await trigger.click()
    const gallery = page.locator('.ribbon-gallery-pane')
    const choices = gallery.locator('.rb-menu button')
    await expect(choices).toHaveCount(13)
    await expect(choices.first()).toBeFocused()
    await assertReservedSpace(page)
    await expect(page.locator('.ai-dock')).toHaveClass(/collapsed/)
    await page.keyboard.press('ArrowRight')
    await expect(choices.nth(1)).toBeFocused()
    await page.keyboard.press('ArrowDown')
    await expect(choices.nth(3)).toBeFocused()
    await page.keyboard.press('End')
    await expect(choices.last()).toBeFocused()
    await expect(choices.last()).toBeInViewport()
    await expect(gallery.locator('header')).toBeInViewport()
    // The pane stays available while the user works in the document.
    await page.locator('.stage-wrap').click({ position: { x: 8, y: 8 } })
    await expect(gallery).toBeVisible()
    await choices.first().focus()
    await page.keyboard.press('Escape')
    await expect(gallery).toHaveCount(0)
    await expect(trigger).toBeFocused()
    expect(await page.evaluate(() => localStorage.getItem('ai-slides-show-ai'))).toBe(aiPreference)
    await trigger.click()
    await choices.first().click()
    await expect(gallery).toHaveCount(0)
    await expect.poll(() => nodes(page)).toBe(before + 1)
    await page.keyboard.press('ControlOrMeta+z')
    await expect.poll(() => nodes(page)).toBe(before)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect.poll(() => nodes(page)).toBe(before + 1)
  } finally {
    await closeAndSaveVideo(launched, 'slides-insert-gallery-chart')
  }
})

test('large galleries replace one another and preserve insertion actions', async () => {
  const { launched, page } = await openSlides()
  try {
    const pane = page.locator('.ribbon-gallery-pane')
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await page.evaluate(() => document.documentElement.setAttribute('data-theme', 'dark'))
    for (const key of ['chart', 'smartart', 'shapes', 'icons', 'wordart']) {
      await page.locator(`[data-gallery-trigger="${key}"]`).click()
      await expect(pane).toHaveCount(1)
      await expect(page.locator('.ribbon-gallery-host')).toHaveAttribute('data-gallery', key)
      await assertReservedSpace(page)
      await page.screenshot({ path: test.info().outputPath(`gallery-${key}.png`) })
    }
    await pane.locator('header button').click()
    await expect(page.locator('[data-gallery-trigger="wordart"]')).toBeFocused()
    for (const [key, selector] of [
      ['smartart', '.rb-menu button'],
      ['icons', '.rb-icon-cell'],
      ['wordart', '.rb-wordart-cell'],
    ]) {
      await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
      const before = await nodes(page)
      await page.locator(`[data-gallery-trigger="${key}"]`).click()
      await pane.locator(selector).first().click()
      await expect(pane).toHaveCount(0)
      await expect.poll(() => nodes(page)).toBeGreaterThan(before)
      await page.keyboard.press('Escape')
    }
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await page.locator('[data-gallery-trigger="chart"]').click()
    await page.locator('.ribbon-tab', { hasText: /^Home$/ }).click()
    await expect(page.locator('.ribbon-gallery-host')).toHaveCount(0)
    await page.locator('.ribbon-tab', { hasText: /^View$/ }).click()
    await page.locator('.ribbon-body button', { hasText: /^Slide Sorter$/ }).click()
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await page.locator('[data-gallery-trigger="chart"]').click()
    await expect(pane).toBeVisible()
    await assertReservedSpace(page)
  } finally {
    await closeAndSaveVideo(launched, 'slides-insert-gallery-actions')
  }
})

test('folded Home Shapes gallery remains usable outside its toolbar menu', async () => {
  const { launched, page } = await openSlides()
  try {
    const drawing = page.locator('[data-rbgroup="drawing"]')
    const trigger = drawing.locator('[data-gallery-trigger="shapes"]')
    if (!(await trigger.isVisible())) await drawing.locator('button.rb-big').first().click()
    await trigger.click()
    await expect(page.locator('.ribbon-gallery-pane')).toBeVisible()
    await page.locator('.stage-wrap').click({ position: { x: 8, y: 8 } })
    await expect(page.locator('.ribbon-gallery-pane')).toBeVisible()
    const before = await nodes(page)
    await page.locator('.ribbon-gallery-pane .rb-shape-cell[data-tip="Rectangle"]').click()
    await expect(page.locator('.ribbon-gallery-pane')).toHaveCount(0)
    const stage = await page.locator('.stage-wrap').boundingBox()
    await page.mouse.move(stage!.x + stage!.width * 0.35, stage!.y + stage!.height * 0.35)
    await page.mouse.down()
    await page.mouse.move(stage!.x + stage!.width * 0.6, stage!.y + stage!.height * 0.6, {
      steps: 8,
    })
    await page.mouse.up()
    await expect.poll(() => nodes(page)).toBe(before + 1)
  } finally {
    await closeAndSaveVideo(launched, 'slides-insert-gallery-folded')
  }
})

test('Traditional Chinese galleries remain accessible in both themes and a narrow window', async () => {
  const testInfo = test.info()
  const { launched, page } = await openSlides('zh-TW')
  try {
    await page.locator('.ribbon-tab', { hasText: /^插入$/ }).click()
    await page.locator('[data-gallery-trigger="chart"]').click()
    for (const theme of ['light', 'dark']) {
      await page.evaluate(
        (value) => document.documentElement.setAttribute('data-theme', value),
        theme,
      )
      await assertReservedSpace(page)
      await page.screenshot({ path: testInfo.outputPath(`gallery-${theme}.png`) })
    }
    await launched.app.evaluate(({ BrowserWindow }) => {
      // Leave room for Windows fractional-DPI rounding at the responsive breakpoint.
      BrowserWindow.getAllWindows()[0].setContentSize(880, 700)
    })
    await expect.poll(() => page.evaluate(() => innerWidth)).toBeLessThanOrEqual(900)
    await assertReservedSpace(page)
    await page.locator('.ribbon-gallery-body .rb-menu button').last().focus()
    await expect(page.locator('.ribbon-gallery-body .rb-menu button').last()).toBeInViewport()
    await page.screenshot({ path: testInfo.outputPath('gallery-narrow.png') })
  } finally {
    await closeAndSaveVideo(launched, 'slides-insert-gallery-responsive')
  }
})
