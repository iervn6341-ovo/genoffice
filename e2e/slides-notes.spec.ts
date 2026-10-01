import { closeAndSaveVideo } from './helpers'
import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

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
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')

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
  test('notes ignore black/white picks and presenter notes stay white on dark in every theme', async () => {
    const { s, launched } = await blankDeck()
    let reopened: LaunchedApp | undefined
    try {
      await s.evaluate(async () => {
        await (window as any).slidesApi.setNotes({
          slideIndex: 0,
          text: 'Black White Red',
          paragraphs: [
            {
              runs: [
                { text: 'Black ', color: '#000000', fontSize: 18 },
                { text: 'White ', color: '#FFFFFF', fontSize: 18 },
                { text: 'Red', color: '#FF0000', fontSize: 18 },
                { text: ' Accent', color: '#C00000', fontSize: 18 },
              ],
            },
          ],
        })
      })
      const file = join(await mkdtemp(join(tmpdir(), 'genoffice-notes-colors-')), 'notes.pptx')
      await launched.app.evaluate(({ dialog }, path) => {
        dialog.showSaveDialog = (async () => ({ canceled: false, filePath: path })) as never
      }, file)
      expect(await s.evaluate(() => (window as any).slidesApi.saveAs('notes.pptx'))).toMatchObject({
        ok: true,
      })
      reopened = await launchShell({
        onboardingSeen: true,
        openFile: file,
        videoDir: 'notes-colors',
      })
      const page = await waitForPageWithUrl(reopened.app, '://slides/', 20000)
      await setEditorLayoutWidth(reopened.app, '://slides/', 2200)
      await expect(notes(page)).toContainText('Black White Red')
      for (const theme of ['dark', 'light']) {
        await page.evaluate(
          (value) => document.documentElement.setAttribute('data-theme', value),
          theme,
        )
        const colors = await notes(page).evaluate((root) => ({
          foreground: getComputedStyle(root).color,
          runs: Array.from(root.querySelectorAll('span')).map((span) => ({
            color: getComputedStyle(span).webkitTextFillColor,
            authored: span.style.color,
          })),
        }))
        expect(colors.runs[0].color).toBe(colors.foreground)
        expect(colors.runs[1].color).toBe(colors.foreground)
        expect(colors.runs[2].color).toBe('rgb(255, 0, 0)')
        expect(colors.runs[0].authored).toBe('rgb(0, 0, 0)')
        expect(colors.runs[1].authored).toBe('rgb(255, 255, 255)')
      }
      expect(await page.evaluate(() => (window as any).slidesApi.isDirty())).toBe(false)
      await notes(page).click()
      await notes(page).evaluate((root) => {
        const range = document.createRange()
        range.selectNodeContents(root.querySelectorAll('span')[2])
        const selection = window.getSelection()!
        selection.removeAllRanges()
        selection.addRange(range)
      })
      await btn(page, 'Font Color').click()
      await expect(page.locator('.rb-color-pop button[title="Black"]')).toBeDisabled()
      await expect(page.locator('.rb-color-pop button[title="White"]')).toBeDisabled()
      const custom = page.locator('.rb-color-pop input[type="color"]')
      const beforePick = await notes(page).innerHTML()
      for (const color of ['#ffffff', '#000000']) {
        await custom.evaluate((input: HTMLInputElement, value) => {
          Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(
            input,
            value,
          )
          input.dispatchEvent(new Event('change', { bubbles: true }))
        }, color)
        // The native picker is debounced by 200ms; wait past its apply callback.
        await page.waitForTimeout(350)
        expect(await notes(page).innerHTML()).toBe(beforePick)
        expect(await page.evaluate(() => (window as any).slidesApi.isDirty())).toBe(false)
        await expect(notes(page).locator('span', { hasText: /^Red$/ })).toHaveCSS(
          'color',
          'rgb(255, 0, 0)',
        )
      }
      await page.getByRole('button', { name: 'Default (follow theme)', exact: true }).click()
      await leaveNotes(page)
      expect(await runOf(page, 'Black')).toMatchObject({ color: '#000000' })
      expect(await runOf(page, 'White')).toMatchObject({ color: '#FFFFFF' })
      expect(await runOf(page, 'Red')).toMatchObject({ color: '#000000' })
      expect(await page.evaluate(() => (window as any).slidesApi.save())).toMatchObject({
        ok: true,
      })
      const savedZip = await JSZip.loadAsync(await readFile(file))
      const notesXml = await savedZip.file('ppt/notesSlides/notesSlide1.xml')!.async('string')
      expect(notesXml).toContain('val="FFFFFF"')
      expect(notesXml).toContain('val="000000"')
      expect(notesXml).toContain('val="C00000"')
      // Force the normal show, then explicitly select presenter view (works on one monitor).
      await page.keyboard.press('F5')
      await page.locator('.slideshow, .presenter').first().waitFor()
      if (!(await page.locator('.pv-notes').count())) {
        await page.locator('.ss-controls button').last().click({ force: true })
        await page.getByText('Use Presenter View', { exact: true }).click()
      }
      await expect(page.locator('.pv-notes')).toContainText('Black White Red')
      for (const theme of ['dark', 'light']) {
        await page.evaluate(
          (value) => document.documentElement.setAttribute('data-theme', value),
          theme,
        )
        const painted = await page.locator('.pv-notes').evaluate((root) => ({
          ink: getComputedStyle(root).color,
          paper: getComputedStyle(root).backgroundColor,
          runs: Array.from(root.querySelectorAll('span')).map(
            (span) => getComputedStyle(span).color,
          ),
        }))
        expect(painted.ink).toBe('rgb(255, 255, 255)')
        expect(painted.paper).toBe('rgb(0, 0, 0)')
        expect(painted.runs.slice(0, -1).every((color) => color === painted.ink)).toBe(true)
        expect(painted.runs.at(-1)).toBe('rgb(192, 0, 0)')
        await page.screenshot({ path: test.info().outputPath(`notes-${theme}.png`) })
      }
      await page.keyboard.press('Escape')
    } finally {
      if (reopened) await kill(reopened)
      await kill(launched)
    }
  })

  test('ribbon Bold / size / font / colour act on the notes selection only', async () => {
    const { s, launched } = await blankDeck()
    let reopened: LaunchedApp | undefined
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
      // The editor shows the authored formatting directly, including after a fresh launch.
      const savedWorld = await runOf(s, 'world')
      const displayed = await notes(s).evaluate((root) => {
        const span = Array.from(root.querySelectorAll('span')).find((el) =>
          el.textContent?.includes('world'),
        )!
        const css = getComputedStyle(span)
        return {
          color: css.color,
          inlineColor: span.style.color,
          size: css.fontSize,
          font: css.fontFamily,
        }
      })
      expect(displayed.color).toBe(displayed.inlineColor)
      expect(displayed.size).toBe('32px')
      expect(displayed.font).toContain('Georgia')
      await expect(
        s.getByRole('button', { name: 'Show original formatting', exact: true }),
      ).toHaveCount(0)
      // the slide itself is untouched
      const slideText = await s.evaluate(async () => {
        const r = await (window as any).slidesApi.getRenderSlides()
        return JSON.stringify(r[0].nodes)
      })
      expect(slideText).not.toContain('world')
      const output =
        process.env.E2E_NOTES_OUTPUT ??
        join(await mkdtemp(join(tmpdir(), 'genoffice-notes-format-')), 'notes.pptx')
      await launched.app.evaluate(({ dialog }, file) => {
        dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
      }, output)
      expect(await s.evaluate(() => (window as any).slidesApi.saveAs('notes.pptx'))).toMatchObject({
        ok: true,
      })
      await writeFile(
        output + '.expected.json',
        JSON.stringify({ changed: savedWorld, untouched: await runOf(s, 'keep') }, null, 2),
      )
      reopened = await launchShell({
        onboardingSeen: true,
        openFile: output,
        videoDir: 'slides-notes-reopen',
      })
      const next = await waitForPageWithUrl(reopened.app, '://slides/', 20000)
      await expect(notes(next)).toContainText('keep world')
      expect(await runOf(next, 'world')).toEqual(savedWorld)
      const world = notes(next).locator('span', { hasText: 'world' })
      await expect(world).toHaveCSS('font-size', '32px')
      await expect(world).toHaveCSS('color', displayed.color)
      await expect(world).toHaveCSS('font-weight', '700')
      expect(await world.evaluate((el) => getComputedStyle(el).fontFamily)).toContain('Georgia')
    } finally {
      if (reopened) await kill(reopened)
      await kill(launched)
    }
  })

  test('⌘B / ⌘I / ⌘⇧> format the notes selection, and the pane shows it after a slide switch', async () => {
    const { s, launched } = await blankDeck()
    try {
      await typeNotes(s, 'alpha beta', 4)
      await s.keyboard.press('ControlOrMeta+b')
      await s.keyboard.press('ControlOrMeta+i')
      await s.keyboard.press('ControlOrMeta+Shift+>') // 12 → 14 on PowerPoint's ladder
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
      await kill(launched)
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
      await s.keyboard.press('ControlOrMeta+b')
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
      await kill(launched)
    }
  })
})
