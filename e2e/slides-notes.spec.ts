import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Speaker notes behave like PowerPoint's notes pane: the Home ribbon's Font commands and the
 * usual shortcuts format the text selected in the notes, the formatting is saved into the
 * notes page, and a show started with "Use Presenter View" on a two-screen desk plays the
 * slides on the other screen while presenter view shows the formatted notes here.
 */

interface NotesRun {
  text: string
  bold: boolean
  italic: boolean
  fontSizePt: number
  fontFamily?: string
  fontExplicit: boolean
  color?: string
}

async function blankDeck(wide = true): Promise<{ s: Page; launched: LaunchedApp }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-notes' })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  // a fully expanded Home ribbon; the show test keeps the real window scale instead
  if (wide) await setEditorLayoutWidth(launched.app, '://slides/', 2200)
  await s.waitForSelector('.notes-editor')
  return { s, launched }
}

const notes = (s: Page) => s.locator('.notes-editor')
const btn = (s: Page, tip: string) => s.locator(`.ribbon-body button[data-tip="${tip}"]`).first()
const kill = (l: LaunchedApp) => l.app.process().kill('SIGKILL')

/** the saved notes run holding `word` (read from the document, not the pane) */
const runOf = (s: Page, word: string, slide = 0): Promise<NotesRun | undefined> =>
  s.evaluate(
    async ([w, i]) => {
      const paras = await (window as any).slidesApi.getNotesRich(i)
      return paras.flatMap((p: any) => p.runs).find((r: any) => r.text.includes(w))
    },
    [word, slide] as const,
  )

/** type into the notes and select the last `n` characters */
async function typeNotes(s: Page, text: string, selectLast: number): Promise<void> {
  await notes(s).click()
  await s.keyboard.type(text)
  for (let i = 0; i < selectLast; i++) await s.keyboard.press('Shift+ArrowLeft')
}

/** click the slide area: the notes lose focus and are saved */
const leaveNotes = async (s: Page) => {
  await s.locator('.stage-wrap').click({ position: { x: 6, y: 6 } })
  await s.waitForTimeout(300)
}

test.describe('notes pane formatting', () => {
  test('ribbon Bold / size / font / colour act on the notes selection only', async () => {
    const { s, launched } = await blankDeck()
    try {
      await typeNotes(s, 'keep world', 5)
      // the Font group is live while typing in the notes
      await expect(btn(s, 'Bold')).toBeEnabled()
      await btn(s, 'Bold').click()
      await expect(btn(s, 'Bold')).toHaveClass(/\bactive\b/)
      const size = s.locator('input.rb-size-input:not(.rb-font-input)').first()
      await size.click()
      await size.fill('24')
      await size.press('Enter')
      const font = s.locator('input.rb-font-input').first()
      await font.click()
      await font.fill('Georgia')
      await font.press('Enter')
      await btn(s, 'Font Color').click()
      await s.locator('.gcp-standard-row .gcp-swatch').nth(1).click()
      await leaveNotes(s)

      await expect
        .poll(() => runOf(s, 'world'))
        .toMatchObject({ bold: true, fontSizePt: 24, fontFamily: 'Georgia' })
      expect((await runOf(s, 'world'))!.color).toMatch(/^#[0-9A-F]{6}$/i)
      // the rest of the note keeps the notes defaults, nothing baked in
      expect(await runOf(s, 'keep')).toMatchObject({
        bold: false,
        fontSizePt: 12,
        fontExplicit: false,
      })
      expect((await runOf(s, 'keep'))!.color).toBeUndefined()
      // the slide itself is untouched
      const slideText = await s.evaluate(async () => {
        const r = await (window as any).slidesApi.getRenderSlides()
        return JSON.stringify(r[0].nodes)
      })
      expect(slideText).not.toContain('world')
    } finally {
      kill(launched)
    }
  })

  test('⌘B / ⌘I / ⌘⇧> format the notes selection, and the pane shows it after a slide switch', async () => {
    const { s, launched } = await blankDeck()
    try {
      await typeNotes(s, 'alpha beta', 4)
      await s.keyboard.press('Meta+b')
      await s.keyboard.press('Meta+i')
      await s.keyboard.press('Meta+Shift+>') // 12 → 14 on PowerPoint's ladder
      await leaveNotes(s)
      await expect
        .poll(() => runOf(s, 'beta'))
        .toMatchObject({ bold: true, italic: true, fontSizePt: 14 })
      expect(await runOf(s, 'alpha')).toMatchObject({ bold: false, fontSizePt: 12 })

      // another slide has its own (empty) notes; coming back shows the formatting
      await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
      await s.locator('.ribbon-body button', { hasText: 'New Slide' }).first().click()
      await expect(notes(s)).toHaveText('')
      await s.locator('.slide-list .thumb').first().click()
      const beta = notes(s).locator('span', { hasText: 'beta' })
      await expect(beta).toHaveCSS('font-weight', '700')
      await expect(beta).toHaveCSS('font-style', 'italic')
      const px = parseFloat(await beta.evaluate((e) => getComputedStyle(e).fontSize))
      expect(px).toBeCloseTo((14 * 96) / 72, 2)
    } finally {
      kill(launched)
    }
  })
})

test.describe('slide show on two screens', () => {
  test('Use Presenter View: slides on the other screen, formatted notes on this one', async () => {
    const { s, launched } = await blankDeck(false)
    try {
      const displays = await launched.app.evaluate(({ screen }) =>
        screen.getAllDisplays().map((d) => ({ id: d.id, label: d.label })),
      )
      test.skip(displays.length < 2, 'needs a second display')
      await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
      await s.waitForSelector('.slide-text-editor')
      await s.keyboard.type('Quarterly review')
      await s.keyboard.press('Escape')
      await s.keyboard.press('Escape')
      await typeNotes(s, 'Say hello first', 5)
      await s.keyboard.press('Meta+b')
      await leaveNotes(s)

      await s.locator('.ribbon-tab', { hasText: 'Slide Show' }).click()
      await expect(
        s.locator('.ribbon-body .rb-check', { hasText: 'Use Presenter View' }),
      ).toHaveClass(/\bon\b/)
      await s
        .locator('.ribbon-body button', { hasText: /From (Beginning|Start)/ })
        .first()
        .click()

      const audience = await waitForPageWithUrl(launched.app, 'mode=audience', 20_000)
      await audience.waitForSelector('.slideshow', { timeout: 10_000 })
      await s.waitForSelector('.pv-notes', { timeout: 10_000 })

      // presenter view here shows the notes with their formatting
      const first = s.locator('.pv-notes span', { hasText: 'first' })
      await expect(first).toHaveCSS('font-weight', '700')
      const before = parseFloat(await first.evaluate((e) => getComputedStyle(e).fontSize))
      await s.locator('.pv-mini-btn', { hasText: 'A⁺' }).click()
      await expect
        .poll(async () => parseFloat(await first.evaluate((e) => getComputedStyle(e).fontSize)))
        .toBeGreaterThan(before)

      // the next-slide / notes column takes about a third of the screen and its divider drags
      const side = s.locator('.pv-side')
      const vw = await s.evaluate(() => window.innerWidth)
      const w0 = (await side.boundingBox())!.width
      expect(w0 / vw).toBeGreaterThan(0.3)
      const bar = (await s.locator('.pv-splitter').boundingBox())!
      await s.mouse.move(bar.x + 3, bar.y + bar.height / 2)
      await s.mouse.down()
      await s.mouse.move(bar.x - 150, bar.y + bar.height / 2, { steps: 5 })
      await s.mouse.up()
      await expect.poll(async () => (await side.boundingBox())!.width).toBeGreaterThan(w0 + 100)

      // the audience window fills another screen: a real monitor rather than Sidecar
      const where = await launched.app.evaluate(({ BrowserWindow, screen }) => {
        const wins = BrowserWindow.getAllWindows()
        const aud = wins.find((w) => w.webContents.getURL().includes('mode=audience'))!
        const host = wins.find((w) => w !== aud && w.isVisible())!
        const a = screen.getDisplayMatching(aud.getBounds())
        return {
          audience: { id: a.id, label: a.label },
          audienceBounds: aud.getBounds(),
          displayBounds: a.bounds,
          host: screen.getDisplayMatching(host.getBounds()).id,
        }
      })
      expect(where.audience.id).not.toBe(where.host)
      expect(where.audienceBounds).toEqual(where.displayBounds)
      const monitors = displays.filter((d) => d.id !== where.host && !/sidecar/i.test(d.label))
      if (monitors.length) expect(where.audience.label).not.toMatch(/sidecar/i)
      await audience.screenshot({ path: test.info().outputPath('audience.png') })
      await s.screenshot({ path: test.info().outputPath('presenter.png') })
      test.info().annotations.push({ type: 'audience display', description: where.audience.label })

      // the presenter's keys drive the audience; Esc ends both
      await s.keyboard.press('Escape')
      await expect
        .poll(() => launched.app.windows().some((w) => w.url().includes('mode=audience')))
        .toBe(false)
      await expect(s.locator('.pv-notes')).toHaveCount(0)
    } finally {
      kill(launched)
    }
  })
})
