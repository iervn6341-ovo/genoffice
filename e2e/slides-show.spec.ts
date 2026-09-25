import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Slide show follows PowerPoint's keyboard: N / P step, a number then Enter jumps to that slide,
 * B (or .) blanks to black and W (or ,) to white, and a second press brings the slide back.
 * The presenter view (two displays) mirrors those onto the audience window.
 */

async function deckOfThree(): Promise<{ s: Page; launched: LaunchedApp }> {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-show',
    foreground: true,
  })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  await setEditorLayoutWidth(launched.app, '://slides/', 1900)
  const addText = async (text: string) => {
    await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
    await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
    await s.waitForSelector('.slide-text-editor')
    await s.keyboard.type(text)
    await s.keyboard.press('Escape')
    await s.waitForTimeout(250)
  }
  await addText('Slide 1')
  for (const n of ['Slide 2', 'Slide 3']) {
    await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
    await s.locator('.ribbon-body button', { hasText: 'New Slide' }).first().click()
    await s.waitForTimeout(400)
    await addText(n)
  }
  return { s, launched }
}

const startShow = async (s: Page, label: RegExp) => {
  await s.locator('.ribbon-tab', { hasText: 'Slide Show' }).click()
  await s.locator('.ribbon-body button', { hasText: label }).first().click()
}

test.describe('slide show keys', () => {
  test('N / P / number+Enter / B / W behave like PowerPoint', async () => {
    const { s, launched } = await deckOfThree()
    try {
      // From the first slide (the deck is on slide 3 after building it), on this screen only:
      // with a second display connected, "Use Presenter View" would split the show
      await s.locator('.ribbon-tab', { hasText: 'Slide Show' }).click()
      const usePv = s.locator('.ribbon-body .rb-check', { hasText: 'Use Presenter View' })
      await expect(usePv).toHaveClass(/\bon\b/)
      await usePv.click()
      await expect(usePv).not.toHaveClass(/\bon\b/)
      await s
        .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
        .first()
        .click()
      await s.waitForSelector('.ss-counter', { timeout: 10_000 })
      const counter = () => s.locator('.ss-counter').innerText()
      await expect.poll(counter).toBe('1 / 3')

      await s.keyboard.press('n')
      await expect.poll(counter).toBe('2 / 3')
      await s.keyboard.press('p')
      await expect.poll(counter).toBe('1 / 3')

      await s.keyboard.press('3')
      await s.keyboard.press('Enter')
      await expect.poll(counter).toBe('3 / 3')

      // a pending number is edited by Backspace instead of going back a slide
      await s.keyboard.press('2')
      await s.keyboard.press('Backspace')
      await expect.poll(counter).toBe('3 / 3')

      await s.keyboard.press('b')
      await expect(s.locator('.ss-screen-black')).toHaveCount(1)
      await s.keyboard.press('b')
      await expect(s.locator('.ss-screen-black')).toHaveCount(0)

      await s.keyboard.press('w')
      await expect(s.locator('.ss-screen-white')).toHaveCount(1)
      await s.keyboard.press('.') // switch from white to black with the other key
      await expect(s.locator('.ss-screen-black')).toHaveCount(1)
      await expect(s.locator('.ss-screen-white')).toHaveCount(0)

      // navigating brings the slide back
      await s.keyboard.press('p')
      await expect(s.locator('.ss-screen')).toHaveCount(0)
      await expect.poll(counter).toBe('2 / 3')

      await s.keyboard.press('Escape')
      await expect(s.locator('.slideshow')).toHaveCount(0)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })
})

const externalDisplayCount = (app: ElectronApplication) =>
  app.evaluate(({ screen }) => screen.getAllDisplays().length)

test.describe('presenter view mirrors the keys onto the audience window', () => {
  test('W / B / number+Enter on the presenter, N from the audience window', async () => {
    const { s, launched } = await deckOfThree()
    try {
      test.skip((await externalDisplayCount(launched.app)) < 2, 'needs a second display')
      await startShow(s, /Presenter View/)
      const audience = await waitForPageWithUrl(launched.app, 'mode=audience', 20_000)
      await audience.waitForSelector('.slideshow', { timeout: 10_000 })
      const pos = () => s.locator('.pv-nav-label').innerText()
      await expect.poll(pos).toMatch(/1\s*\D+\s*3|1 of 3/)

      await s.keyboard.press('w')
      await expect(audience.locator('.ss-white')).toHaveCount(1)
      await s.keyboard.press('b')
      await expect(audience.locator('.ss-black')).toHaveCount(1)
      await expect(audience.locator('.ss-white')).toHaveCount(0)
      await s.keyboard.press('b')
      await expect(audience.locator('.ss-black')).toHaveCount(0)

      await s.keyboard.press('3')
      await s.keyboard.press('Enter')
      await expect.poll(pos).toMatch(/3/)

      // keys pressed on the projector screen drive the presenter too
      await audience.keyboard.press('p')
      await expect.poll(pos).toMatch(/2/)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })
})
