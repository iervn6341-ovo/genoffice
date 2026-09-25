import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Home → Font styles, driven through the real ribbon and keyboard in both PowerPoint modes:
 * a selected text box (the command applies to all of its text) and typing in the box (the
 * command applies to the selection and survives the commit). Results are read back from the
 * rendered slide model, and the B / I / U / S / x² / x₂ buttons must light up like Office's.
 */

interface Run {
  text: string
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  baseline: number
  sizePx: number
  family: string
  color: string
}

async function deckWithBox(text: string): Promise<{ s: Page; launched: LaunchedApp }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-font-styles' })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  // a fully expanded Home ribbon, in a window no bigger than TEST_WINDOW
  await setEditorLayoutWidth(launched.app, '://slides/', 2200)
  await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
  await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
  await s.waitForSelector('.slide-text-editor')
  await s.keyboard.type(text)
  await s.keyboard.press('Escape') // commit, box stays selected
  await s.waitForTimeout(400)
  await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
  return { s, launched }
}

const runs = (s: Page): Promise<Run[]> =>
  s.evaluate(async () => {
    const r = await (window as any).slidesApi.getRenderSlides()
    const n = r[0].nodes.find((x: any) => x.text)
    return n.text.lines
      .flatMap((l: any) => l.runs)
      .filter((x: any) => !x.isBullet && x.text.trim())
      .map((x: any) => ({
        text: x.text,
        bold: !!x.bold,
        italic: !!x.italic,
        underline: !!x.underline,
        strike: !!x.strike,
        baseline: x.baselinePct ?? 0,
        sizePx: x.fontSizePx,
        family: String(x.srcFontFamily ?? x.fontFamily),
        color: String(x.color).toUpperCase(),
      }))
  })

/** the run holding `word` */
const runOf = async (s: Page, word: string) => (await runs(s)).find((r) => r.text.includes(word))!

const btn = (s: Page, tip: string) => s.locator(`.ribbon-body button[data-tip="${tip}"]`).first()
const pressed = (s: Page, tip: string) =>
  btn(s, tip).evaluate((b) => b.classList.contains('active'))

const kill = (l: LaunchedApp) => l.app.process().kill('SIGKILL')

test.describe('font styles on a selected text box (whole box)', () => {
  test('B / I / U / S toggle on and off, light up, and undo', async () => {
    const { s, launched } = await deckWithBox('Styled text')
    try {
      for (const [tip, key] of [
        ['Bold', 'bold'],
        ['Italic', 'italic'],
        ['Underline', 'underline'],
        ['Strikethrough', 'strike'],
      ] as const) {
        await btn(s, tip).click()
        await expect.poll(async () => (await runs(s)).every((r) => r[key])).toBe(true)
        await expect.poll(() => pressed(s, tip), `${tip} lights up`).toBe(true)
        await btn(s, tip).click()
        await expect.poll(async () => (await runs(s)).some((r) => r[key])).toBe(false)
        await expect.poll(() => pressed(s, tip)).toBe(false)
      }
      await btn(s, 'Bold').click()
      await expect.poll(async () => (await runs(s))[0]!.bold).toBe(true)
      await s.keyboard.press('Meta+z')
      await expect.poll(async () => (await runs(s))[0]!.bold).toBe(false)
    } finally {
      kill(launched)
    }
  })

  test('⌘B / ⌘I / ⌘U work without entering the box', async () => {
    const { s, launched } = await deckWithBox('Keys')
    try {
      await s.keyboard.press('Meta+b')
      await expect.poll(async () => (await runs(s))[0]!.bold).toBe(true)
      await s.keyboard.press('Meta+i')
      await expect.poll(async () => (await runs(s))[0]!.italic).toBe(true)
      await s.keyboard.press('Meta+u')
      await expect.poll(async () => (await runs(s))[0]!.underline).toBe(true)
      await s.keyboard.press('Meta+b')
      await expect.poll(async () => (await runs(s))[0]!.bold).toBe(false)
    } finally {
      kill(launched)
    }
  })

  test('superscript / subscript toggle and exclude each other', async () => {
    const { s, launched } = await deckWithBox('x2')
    try {
      await btn(s, 'Superscript').click()
      await expect.poll(async () => (await runs(s))[0]!.baseline).toBeGreaterThan(0)
      await expect.poll(() => pressed(s, 'Superscript')).toBe(true)
      await btn(s, 'Subscript').click()
      await expect.poll(async () => (await runs(s))[0]!.baseline).toBeLessThan(0)
      await expect.poll(() => pressed(s, 'Superscript')).toBe(false)
      await btn(s, 'Subscript').click()
      await expect.poll(async () => (await runs(s))[0]!.baseline).toBe(0)
    } finally {
      kill(launched)
    }
  })

  test('font size box, grow / shrink along the ladder, and ⌘⇧> / ⌘⇧<', async () => {
    const { s, launched } = await deckWithBox('Size')
    try {
      const base = (await runs(s))[0]!.sizePx // 18 pt
      const pt = async () => Math.round(((await runs(s))[0]!.sizePx / base) * 18 * 2) / 2
      const size = s.locator('input.rb-size-input:not(.rb-font-input)').first()
      await size.click()
      await size.fill('36')
      await size.press('Enter')
      await expect.poll(pt).toBe(36)
      await btn(s, 'Increase Font Size').click()
      await expect.poll(pt).toBe(40)
      await btn(s, 'Decrease Font Size').click()
      await btn(s, 'Decrease Font Size').click()
      await expect.poll(pt).toBe(32)
      await s
        .locator('.slide-thumb, .ribbon')
        .first()
        .evaluate(() => (document.activeElement as HTMLElement)?.blur())
      await s.keyboard.press('Meta+Shift+>')
      await expect.poll(pt).toBe(36)
      await s.keyboard.press('Meta+Shift+<')
      await expect.poll(pt).toBe(32)
    } finally {
      kill(launched)
    }
  })

  test('font family and font colour', async () => {
    const { s, launched } = await deckWithBox('Family')
    try {
      const font = s.locator('input.rb-font-input').first()
      await font.click()
      await font.fill('Georgia')
      await font.press('Enter')
      await expect.poll(async () => (await runs(s))[0]!.family).toContain('Georgia')
      await btn(s, 'Font Color').click()
      const swatch = s.locator('.gcp-standard-row .gcp-swatch').nth(1)
      const want = await swatch.evaluate((b) => {
        const m = getComputedStyle(b).backgroundColor.match(/\d+/g)!.map(Number)
        return (
          '#' +
          m
            .slice(0, 3)
            .map((v) => v.toString(16).padStart(2, '0'))
            .join('')
            .toUpperCase()
        )
      })
      await swatch.click()
      await expect.poll(async () => (await runs(s))[0]!.color).toBe(want)
    } finally {
      kill(launched)
    }
  })
})

test.describe('font styles while typing (selection only, kept after commit)', () => {
  /** enter the box and select the last word (5 letters) */
  const selectLastWord = async (s: Page) => {
    await s.keyboard.press('Enter')
    await s.waitForSelector('.slide-text-editor')
    for (let i = 0; i < 5; i++) await s.keyboard.press('Shift+ArrowLeft')
  }
  const commit = async (s: Page) => {
    await s.keyboard.press('Escape')
    await s.waitForTimeout(600)
  }

  test('ribbon B / I / U / S / x² on the selected word only', async () => {
    const { s, launched } = await deckWithBox('keep world')
    try {
      await selectLastWord(s)
      await btn(s, 'Bold').click()
      await expect.poll(() => pressed(s, 'Bold')).toBe(true)
      await btn(s, 'Italic').click()
      await btn(s, 'Underline').click()
      await btn(s, 'Strikethrough').click()
      await btn(s, 'Superscript').click()
      await expect.poll(() => pressed(s, 'Superscript')).toBe(true)
      expect(await s.evaluate(() => window.getSelection()?.toString())).toBe('world')
      await commit(s)
      const w = await runOf(s, 'world')
      expect(w).toMatchObject({ bold: true, italic: true, underline: true, strike: true })
      expect(w.baseline).toBeGreaterThan(0)
      const k = await runOf(s, 'keep')
      expect(k).toMatchObject({ bold: false, italic: false, underline: false, strike: false })
      expect(k.baseline).toBe(0)
    } finally {
      kill(launched)
    }
  })

  test('⌘B / ⌘I / ⌘U inside the box light the buttons', async () => {
    const { s, launched } = await deckWithBox('keep world')
    try {
      await selectLastWord(s)
      await s.keyboard.press('Meta+b')
      await s.keyboard.press('Meta+i')
      await s.keyboard.press('Meta+u')
      await expect.poll(() => pressed(s, 'Bold')).toBe(true)
      await expect.poll(() => pressed(s, 'Italic')).toBe(true)
      await expect.poll(() => pressed(s, 'Underline')).toBe(true)
      await commit(s)
      expect(await runOf(s, 'world')).toMatchObject({ bold: true, italic: true, underline: true })
      expect((await runOf(s, 'keep')).bold).toBe(false)
    } finally {
      kill(launched)
    }
  })

  test('size, grow / shrink, ⌘⇧>, family and colour on the selection', async () => {
    const { s, launched } = await deckWithBox('keep world')
    try {
      const base = (await runs(s))[0]!.sizePx
      await selectLastWord(s)
      await btn(s, 'Increase Font Size').click() // 18 → 20
      await s.keyboard.press('Meta+Shift+>') // 20 → 24
      await btn(s, 'Decrease Font Size').click() // 24 → 20
      await s.keyboard.press('Meta+Shift+>') // 20 → 24
      const font = s.locator('input.rb-font-input').first()
      await font.click()
      await font.fill('Georgia')
      await font.press('Enter')
      await btn(s, 'Font Color').click()
      await s.locator('.gcp-standard-row .gcp-swatch').nth(1).click()
      await commit(s)
      const w = await runOf(s, 'world')
      const k = await runOf(s, 'keep')
      expect(Math.round((w.sizePx / base) * 18)).toBe(24)
      expect(Math.round((k.sizePx / base) * 18)).toBe(18)
      expect(w.family).toContain('Georgia')
      expect(k.family).not.toContain('Georgia')
      expect(w.color).not.toBe(k.color)
    } finally {
      kill(launched)
    }
  })

  test('Clear All Formatting removes the styles from the selection', async () => {
    const { s, launched } = await deckWithBox('keep world')
    try {
      await selectLastWord(s)
      await btn(s, 'Bold').click()
      await btn(s, 'Italic').click()
      await btn(s, 'Clear All Formatting').click()
      await commit(s)
      expect(await runOf(s, 'world')).toMatchObject({ bold: false, italic: false })
    } finally {
      kill(launched)
    }
  })
})
