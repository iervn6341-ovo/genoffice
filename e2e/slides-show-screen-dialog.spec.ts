import { expect, test } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Starting a show with presenter view while another screen is connected always asks which screen
 * the audience sees, and starts on the screen picked last time. Needs the real displays, so it
 * skips on fewer than three (it picks between two external screens).
 */
test('the show asks which external screen to use before it starts', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'slides-show-screen-dialog',
    foreground: true,
  })
  try {
    const count = await launched.app.evaluate(({ screen }) => screen.getAllDisplays().length)
    test.skip(count < 3, 'needs a laptop screen and two external screens')
    await launched.page.locator('.quick-card').nth(2).click()
    const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
    await s.waitForSelector('.ribbon', { timeout: 15_000 })
    await s.evaluate(() => {
      localStorage.removeItem('genoffice.slides.showMonitor')
    })
    await s.locator('.ribbon-tab', { hasText: 'Slide Show' }).click()
    const usePv = s.locator('.ribbon-body .rb-check', { hasText: 'Use Presenter View' })
    if (!(await usePv.evaluate((b) => b.classList.contains('on')))) await usePv.click()
    const fromStart = s
      .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
      .first()
    const showDisplay = () =>
      launched.app.evaluate(({ BrowserWindow, screen }) => {
        const w = BrowserWindow.getAllWindows().find((x) =>
          x.webContents.getURL().includes('mode=audience'),
        )
        return w ? screen.getDisplayMatching(w.getBounds()).id : null
      })

    // 1. a dialog first: the presenter's own screen is not offered, nothing has started yet
    await fromStart.click()
    const dialog = s.locator('.ssd')
    await expect(dialog).toBeVisible()
    const screens = await s.evaluate(() => window.slidesApi.displayList())
    const external = screens.filter((x) => !x.current)
    expect(external.length).toBeGreaterThanOrEqual(2)
    await expect(dialog.locator('.ssd-item')).toHaveCount(external.length)
    await expect(dialog.getByText(screens.find((x) => x.current)!.label).first()).toBeVisible()
    await expect(s.locator('.presenter')).toHaveCount(0)

    // Cancel starts nothing
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(dialog).toHaveCount(0)
    await expect(s.locator('.presenter')).toHaveCount(0)

    // 2. pick the second external screen: the show opens there
    const target = external[1]!
    await fromStart.click()
    await dialog.locator('.ssd-item', { hasText: target.label }).click()
    await dialog.getByRole('button', { name: 'Start Show' }).click()
    await s.waitForSelector('.pv-stagebox', { timeout: 20_000 })
    await waitForPageWithUrl(launched.app, 'mode=audience', 20_000)
    await expect.poll(showDisplay).toBe(target.id)
    await s.keyboard.press('Escape')
    await expect(s.locator('.presenter')).toHaveCount(0)

    // 3. every start asks again, starting on the screen picked last time
    await fromStart.click()
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('.ssd-item.active')).toContainText(target.label)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(s.locator('.presenter')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'cleanup')
  }
})
