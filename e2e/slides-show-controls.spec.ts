import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * The show controls follow PowerPoint for Mac: Pointer Options (laser pointer, pen,
 * highlighter, eraser, erase all ink, ink colours), Zoom, Black Screen and the "…" Slide Show
 * Options menu, in presenter view and on the full-screen show's bottom-left bar. On two
 * screens the laser is mirrored every frame and the keyboard stays with the presenter, so Esc
 * ends the show.
 */

async function deckOfThree(chromiumArgs?: string[]): Promise<{ s: Page; launched: LaunchedApp }> {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-show-controls',
    // full-screen shows and the presenter keyboard need the real foreground
    foreground: true,
    ...(chromiumArgs ? { chromiumArgs } : {}),
  })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  for (const [i, title] of ['Opening', 'Results', 'Next steps'].entries()) {
    if (i > 0) {
      await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
      await s.locator('.ribbon-body button', { hasText: 'New Slide' }).first().click()
      await s.waitForTimeout(300)
    }
    await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
    await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
    await s.waitForSelector('.slide-text-editor')
    await s.keyboard.type(title)
    await s.keyboard.press('Escape')
    await s.keyboard.press('Escape')
  }
  await s.locator('.ribbon-tab', { hasText: 'Slide Show' }).click()
  return { s, launched }
}

const displayCount = (l: LaunchedApp) =>
  l.app.evaluate(({ screen }) => screen.getAllDisplays().length)

const bar = (p: Page, label: string) => p.locator(`.ssc-btn[aria-label="${label}"]`).first()

/** Opaque pixels on a page's ink canvas */
const inkPixels = (p: Page) =>
  p.evaluate(() => {
    const c = document.querySelector('.ink-layer canvas') as HTMLCanvasElement | null
    if (!c) return -1
    const d = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data
    let n = 0
    for (let i = 3; i < d.length; i += 4) if (d[i]! > 0) n++
    return n
  })

async function drag(p: Page, box: { x: number; y: number; width: number; height: number }) {
  await p.mouse.move(box.x + box.width * 0.2, box.y + box.height * 0.5)
  await p.mouse.down()
  await p.mouse.move(box.x + box.width * 0.6, box.y + box.height * 0.55, { steps: 12 })
  await p.mouse.up()
}

test.describe('presenter view controls (two screens)', () => {
  test('laser is mirrored every frame, Esc ends the show with real keyboard focus', async () => {
    const { s, launched } = await deckOfThree()
    try {
      test.skip((await displayCount(launched)) < 2, 'needs a second display')
      await s
        .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
        .first()
        .click()
      const audience = await waitForPageWithUrl(launched.app, 'mode=audience', 20_000)
      await audience.waitForSelector('.slideshow .ss-stagebox')
      await s.waitForSelector('.pv-stagebox')
      await s.waitForTimeout(300)

      // whichever show window is key, a page in it holds the keyboard (before, the audience
      // window became key with no focused page, so real key presses went nowhere)
      await launched.app.evaluate(({ app }) => app.focus({ steal: true }))
      const focused = () =>
        launched.app.evaluate(
          ({ webContents }) => webContents.getFocusedWebContents()?.getURL() ?? '',
        )
      await expect.poll(focused).toContain('://slides/')

      await bar(s, 'Pointer Options').click()
      await expect(s.locator('.ssc-pop .ssc-pop-item')).toHaveText([
        'Laser Pointer',
        'Pen',
        'Highlighter',
        'Eraser',
        'Erase All Ink on Slide',
      ])
      await expect(s.locator('.ssc-pop .ssc-swatch')).toHaveCount(12)
      await s.locator('.ssc-pop-item', { hasText: 'Laser Pointer' }).click()

      await audience.evaluate(() => {
        const dot = document.querySelector('.ink-laser') as HTMLElement
        ;(window as any).__moves = 0
        new MutationObserver(() => (window as any).__moves++).observe(dot, {
          attributes: true,
          attributeFilter: ['style'],
        })
      })
      const box = (await s.locator('.pv-stagebox').boundingBox())!
      const moves = 40
      for (let i = 0; i < moves; i++) {
        await s.mouse.move(box.x + box.width * (0.2 + i * 0.015), box.y + box.height * 0.4)
        await s.waitForTimeout(20)
      }
      // one update per display frame reaches the audience (the old 30 ms throttle plus a
      // full re-render per move dropped most of them)
      await expect
        .poll(() => audience.evaluate(() => (window as any).__moves))
        .toBeGreaterThan(moves * 0.6)
      const at = await audience.evaluate(() => {
        const dot = document.querySelector('.ink-laser') as HTMLElement
        const frame = document.querySelector('.ss-stagebox') as HTMLElement
        const m = /translate\(([\d.]+)px, ([\d.]+)px\)/.exec(dot.style.transform)!
        return { x: Number(m[1]) / frame.clientWidth, y: Number(m[2]) / frame.clientHeight }
      })
      expect(at.x).toBeCloseTo(0.2 + (moves - 1) * 0.015, 1)
      expect(at.y).toBeCloseTo(0.4, 1)

      // Esc on the page that holds the keyboard ends the show at once, laser still on
      const keyPage = (await focused()).includes('mode=audience') ? audience : s
      await keyPage.keyboard.press('Escape')
      await expect(s.locator('.presenter')).toHaveCount(0)
      await expect
        .poll(() => launched.app.windows().some((w) => w.url().includes('mode=audience')))
        .toBe(false)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('pen, eraser, E and ink colours reach the audience; the … menu works', async () => {
    const { s, launched } = await deckOfThree()
    try {
      test.skip((await displayCount(launched)) < 2, 'needs a second display')
      await s
        .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
        .first()
        .click()
      const audience = await waitForPageWithUrl(launched.app, 'mode=audience', 20_000)
      await audience.waitForSelector('.slideshow .ss-stagebox')
      const stage = s.locator('.pv-stagebox')
      await stage.waitFor()

      // a colour picks the pen (PowerPoint)
      await bar(s, 'Pointer Options').click()
      await s.locator('.ssc-swatch[aria-label="Dark Blue"]').click()
      await expect(bar(s, 'Pointer Options')).toHaveClass(/ssc-on/)
      await drag(s, (await stage.boundingBox())!)
      await expect.poll(() => inkPixels(audience)).toBeGreaterThan(50)
      expect(await inkPixels(s)).toBeGreaterThan(50)

      // eraser across the stroke removes it on both screens
      await bar(s, 'Pointer Options').click()
      await s.locator('.ssc-pop-item', { hasText: 'Eraser' }).click()
      await drag(s, (await stage.boundingBox())!)
      await expect.poll(() => inkPixels(audience)).toBe(0)

      // highlighter, then E erases all ink
      await s.keyboard.press('Control+i')
      await drag(s, (await stage.boundingBox())!)
      await expect.poll(() => inkPixels(audience)).toBeGreaterThan(50)
      await s.keyboard.press('e')
      await expect.poll(() => inkPixels(audience)).toBe(0)
      await s.keyboard.press('Control+a') // back to the arrow

      // the … menu has PowerPoint's items
      await bar(s, 'Slide Show Options').click()
      const items = s.locator('.ssc-more > .ssc-menu-row > .ssc-menu-item .ssc-menu-label')
      await expect(items).toHaveText([
        'Next',
        'Previous',
        'Last Viewed',
        'By Title',
        'Screen',
        'Swap Displays',
        'Use Slide Show',
        'Pointer Options',
        'Pause',
        'End Show',
      ])
      // By Title jumps to a slide by its title
      await s.locator('.ssc-menu-item', { hasText: 'By Title' }).hover()
      await s.locator('.ssc-submenu .ssc-menu-item', { hasText: 'Next steps' }).click()
      await expect(s.locator('.pv-nav-label')).toContainText('3')
      // Screen > Black Screen blacks out the audience screen
      await bar(s, 'Slide Show Options').click()
      await s.locator('.ssc-menu-item', { hasText: 'Screen' }).hover()
      await s.locator('.ssc-submenu .ssc-menu-item', { hasText: 'Black Screen' }).click()
      await expect(audience.locator('.ss-black')).toHaveCount(1)
      // Esc closes an open menu first, then ends the show
      await bar(s, 'Slide Show Options').click()
      await s.keyboard.press('Escape')
      await expect(s.locator('.ssc-more')).toHaveCount(0)
      await expect(s.locator('.presenter')).toHaveCount(1)
      await s.keyboard.press('Escape')
      await expect(s.locator('.presenter')).toHaveCount(0)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })
})

test('full-screen show: bottom-left control bar, pen, zoom and Esc order', async () => {
  const { s, launched } = await deckOfThree()
  try {
    const usePv = s.locator('.ribbon-body .rb-check', { hasText: 'Use Presenter View' })
    if (await usePv.evaluate((b) => b.classList.contains('on'))) await usePv.click()
    await s
      .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
      .first()
      .click()
    await s.waitForSelector('.slideshow .ss-controls')
    await s.mouse.move(300, 300)
    await s.mouse.move(320, 320)
    await expect(s.locator('.ss-controls')).toHaveClass(/ss-controls-on/)
    // ‹ › pointer zoom black camera …
    await expect(s.locator('.ss-controls .ssc-btn')).toHaveCount(7)

    await bar(s, 'Pointer Options').click()
    await s.locator('.ssc-pop-item', { hasText: 'Pen' }).click()
    const stage = s.locator('.slideshow .ss-frame > div').first()
    await drag(s, (await stage.boundingBox())!)
    await expect.poll(() => inkPixels(s)).toBeGreaterThan(50)
    await expect(s.locator('.ss-counter')).toHaveText('1 / 3') // drawing does not advance

    // Zoom: arm, click a spot, click again to zoom out
    await s.keyboard.press('Control+a')
    await bar(s, 'Zoom into the slide').click()
    const b = (await stage.boundingBox())!
    await s.mouse.click(b.x + b.width * 0.25, b.y + b.height * 0.25)
    await expect(s.locator('.slideshow .show-zoom')).toHaveAttribute('style', /scale\(2\)/)
    await s.mouse.click(b.x + b.width * 0.5, b.y + b.height * 0.5)
    await expect(s.locator('.slideshow .show-zoom')).not.toHaveAttribute('style', /scale/)
    await expect(s.locator('.ss-counter')).toHaveText('1 / 3')

    // the … menu opens upward from the bar; Esc closes it, the next Esc ends the show
    await bar(s, 'Slide Show Options').click()
    await expect(s.locator('.ssc-more')).toBeVisible()
    await expect(
      s.locator('.ssc-more .ssc-menu-item', { hasText: 'Use Presenter View' }),
    ).toHaveCount(1)
    await s.keyboard.press('Escape')
    await expect(s.locator('.ssc-more')).toHaveCount(0)
    await expect(s.locator('.slideshow')).toHaveCount(1)
    await s.keyboard.press('Escape')
    await expect(s.locator('.slideshow')).toHaveCount(0)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

/** Chromium's stand-in camera: tests never open the real one */
const FAKE_CAMERA = ['--use-fake-device-for-media-stream', '--use-fake-ui-for-media-stream']

test('audience screen: picked Monitor, mouse laser and pen mirror back, camera on both screens', async () => {
  const { s, launched } = await deckOfThree(FAKE_CAMERA)
  try {
    const displays = await launched.app.evaluate(({ screen }) =>
      screen
        .getAllDisplays()
        .map((d) => ({ label: d.label, primary: d.id === screen.getPrimaryDisplay().id })),
    )
    test.skip(displays.length < 2, 'needs a second display')
    // Slide Show → Monitor: pick a named screen other than the presenter's (the primary)
    const target = displays.find((d) => !d.primary)!
    const picker = s.locator('.rb-monitor select')
    await expect(picker.locator('option').first()).toHaveText('Automatic')
    await picker.selectOption(target.label)
    await s
      .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
      .first()
      .click()
    const audience = await waitForPageWithUrl(launched.app, 'mode=audience', 20_000)
    await audience.waitForSelector('.slideshow .ss-stagebox')
    await s.waitForSelector('.pv-stagebox')
    const shownOn = await launched.app.evaluate(({ BrowserWindow, screen }) => {
      const aud = BrowserWindow.getAllWindows().find((w) =>
        w.webContents.getURL().includes('mode=audience'),
      )!
      return screen.getDisplayMatching(aud.getBounds()).label
    })
    expect(shownOn).toBe(target.label)

    // laser: moving the mouse on the audience screen shows it there and in presenter view
    await s.keyboard.press('Control+l')
    const aBox = (await audience.locator('.ss-stagebox').boundingBox())!
    await audience.mouse.move(aBox.x + aBox.width * 0.3, aBox.y + aBox.height * 0.6)
    await audience.mouse.move(aBox.x + aBox.width * 0.35, aBox.y + aBox.height * 0.6, { steps: 4 })
    await expect(audience.locator('.ink-laser')).toBeVisible()
    await expect(s.locator('.pv-stagebox .ink-laser')).toBeVisible()
    const pAt = await s.evaluate(() => {
      const dot = document.querySelector('.pv-stagebox .ink-laser') as HTMLElement
      const frame = document.querySelector('.pv-stagebox') as HTMLElement
      const m = /translate\(([\d.]+)px, ([\d.]+)px\)/.exec(dot.style.transform)!
      return { x: Number(m[1]) / frame.clientWidth, y: Number(m[2]) / frame.clientHeight }
    })
    expect(pAt.x).toBeCloseTo(0.35, 1)
    expect(pAt.y).toBeCloseTo(0.6, 1)
    // clicking with a pointer tool does not advance the show
    await audience.mouse.click(aBox.x + aBox.width * 0.5, aBox.y + aBox.height * 0.5)
    await expect(s.locator('.pv-nav-label')).toContainText('1')

    // pen drawn on the audience screen appears in presenter view
    await s.keyboard.press('Control+p')
    await drag(audience, aBox)
    await expect.poll(() => inkPixels(s)).toBeGreaterThan(50)

    // Camera: a live bubble on both screens (Chromium's fake device here)
    await bar(s, 'Camera').click()
    const playing = (p: Page) =>
      p.evaluate(() => {
        const v = document.querySelector('.show-camera video') as HTMLVideoElement | null
        return !!v && v.readyState >= 2 && v.videoWidth > 0
      })
    await expect.poll(() => playing(s)).toBe(true)
    await expect.poll(() => playing(audience)).toBe(true)
    await bar(s, 'Camera').click()
    await expect(audience.locator('.show-camera')).toHaveCount(0)
    await s.keyboard.press('Escape')
    await expect(s.locator('.presenter')).toHaveCount(0)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
