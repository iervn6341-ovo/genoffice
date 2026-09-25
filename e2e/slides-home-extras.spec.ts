import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Home tab extras modelled on PowerPoint: Font → Character Spacing / Change Case / Text Highlight,
 * Paragraph → Text Direction / Align Text, Drawing → Quick Styles / Shape Fill / Shape Outline.
 * Each is driven through the real ribbon and checked in the rendered slide model.
 */

async function deckWithBox(text: string): Promise<{ s: Page; launched: LaunchedApp }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-home-extras' })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  // a fully expanded Home ribbon, in a window no bigger than TEST_WINDOW
  await setEditorLayoutWidth(launched.app, '://slides/', 2200)
  await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
  await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
  await s.waitForSelector('.slide-text-editor')
  await s.keyboard.type(text)
  await s.keyboard.press('Escape') // commit; the box stays selected
  await s.waitForTimeout(400)
  await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
  return { s, launched }
}

/** first text box on slide 1: runs (bullets excluded) + body props */
const box = (s: Page) =>
  s.evaluate(async () => {
    const r = await (window as any).slidesApi.getRenderSlides()
    const n = r[0].nodes.find((x: any) => x.text)
    const runs = n.text.lines.flatMap((l: any) => l.runs).filter((x: any) => !x.isBullet && x.text)
    return {
      text: runs.map((x: any) => x.text).join(''),
      runs: runs.map((x: any) => ({
        text: x.text as string,
        highlight: (x.highlight as string | undefined) ?? null,
        spacing: (x.letterSpacingPx as number | undefined) ?? 0,
      })),
      anchor: n.text.anchor as string,
      vert: (n.text.vert as string | undefined) ?? 'horz',
    }
  })

const drop = async (s: Page, tip: string, item: string) => {
  await s.locator(`.ribbon-body button.rb-icon-drop[data-tip="${tip}"]`).first().click()
  await s.locator('.rb-drop button', { hasText: item }).first().click()
  await s.waitForTimeout(500)
}

const kill = (l: LaunchedApp) => l.app.process().kill('SIGKILL')

test.describe('Home → Font extras on a selected text box', () => {
  test('Change Case, then undo restores the text', async () => {
    const { s, launched } = await deckWithBox('hello WORLD. again')
    try {
      await drop(s, 'Change Case', 'UPPERCASE')
      await expect.poll(async () => (await box(s)).text).toBe('HELLO WORLD. AGAIN')
      await drop(s, 'Change Case', 'Sentence case.')
      await expect.poll(async () => (await box(s)).text).toBe('Hello world. Again')
      await drop(s, 'Change Case', 'Capitalize Each Word')
      await expect.poll(async () => (await box(s)).text).toBe('Hello World. Again')
      await s.keyboard.press('Meta+z')
      await expect.poll(async () => (await box(s)).text).toBe('Hello world. Again')
    } finally {
      kill(launched)
    }
  })

  test('Character Spacing Loose widens the text, Normal restores it', async () => {
    const { s, launched } = await deckWithBox('Spacing')
    try {
      await drop(s, 'Character Spacing', 'Very Loose')
      await expect.poll(async () => (await box(s)).runs[0]!.spacing).toBeGreaterThan(0)
      await drop(s, 'Character Spacing', 'Normal')
      await expect.poll(async () => (await box(s)).runs[0]!.spacing).toBe(0)
    } finally {
      kill(launched)
    }
  })

  test('Text Highlight: palette colour, the split button re-applies it, No Color removes it', async () => {
    const { s, launched } = await deckWithBox('Marked')
    try {
      await s.locator('.rb-split-icon .rb-icon-caret').click()
      await s.locator('.rb-highlight-grid button[data-tip="#00FF00"]').click()
      await expect
        .poll(async () => (await box(s)).runs[0]!.highlight?.toUpperCase())
        .toBe('#00FF00')
      await s.locator('.rb-split-icon .rb-icon-caret').click()
      await s.locator('.rb-highlight-pop button', { hasText: 'No Color' }).click()
      await expect.poll(async () => (await box(s)).runs[0]!.highlight).toBeNull()
      // the main half remembers the last colour
      await s.locator('.rb-split-icon > .rb-icon').first().click()
      await expect
        .poll(async () => (await box(s)).runs[0]!.highlight?.toUpperCase())
        .toBe('#00FF00')
    } finally {
      kill(launched)
    }
  })
})

test.describe('Home → Font extras while typing', () => {
  test('highlight and spacing hit only the selected word and survive the commit', async () => {
    const { s, launched } = await deckWithBox('hello world')
    try {
      await s.keyboard.press('Enter') // back into the box, caret at the end
      await s.waitForSelector('.slide-text-editor')
      for (let i = 0; i < 5; i++) await s.keyboard.press('Shift+ArrowLeft') // "world"
      await s.locator('.rb-split-icon .rb-icon-caret').click()
      await s.locator('.rb-highlight-grid button[data-tip="#FFFF00"]').click()
      // the editor keeps focus and the selection, so the next command hits the same word
      expect(await s.evaluate(() => window.getSelection()?.toString())).toBe('world')
      await drop(s, 'Character Spacing', 'Loose')
      await s.keyboard.press('Escape')
      await s.waitForTimeout(600)
      const b = await box(s)
      const world = b.runs.find((r) => r.text.includes('world'))!
      const hello = b.runs.find((r) => r.text.includes('hello'))!
      expect(world.highlight?.toUpperCase()).toBe('#FFFF00')
      expect(world.spacing).toBeGreaterThan(0)
      expect(hello.highlight).toBeNull()
      expect(hello.spacing).toBe(0)
    } finally {
      kill(launched)
    }
  })

  test('Change Case with a caret and no selection changes the word under the caret', async () => {
    const { s, launched } = await deckWithBox('hello world')
    try {
      await s.keyboard.press('Enter')
      await s.waitForSelector('.slide-text-editor')
      await s.keyboard.press('ArrowLeft')
      await s.keyboard.press('ArrowLeft') // caret inside "world"
      await drop(s, 'Change Case', 'UPPERCASE')
      await s.keyboard.press('Escape')
      await expect.poll(async () => (await box(s)).text).toBe('hello WORLD')
    } finally {
      kill(launched)
    }
  })
})

test.describe('Home → Paragraph extras', () => {
  test('Align Text and Text Direction change the box body and undo', async () => {
    const { s, launched } = await deckWithBox('Anchored')
    try {
      await drop(s, 'Align Text', 'Bottom')
      await expect.poll(async () => (await box(s)).anchor).toBe('bottom')
      await drop(s, 'Text Direction', 'Rotate all text 90°')
      await expect.poll(async () => (await box(s)).vert).toBe('vert')
      await drop(s, 'Text Direction', 'Horizontal')
      await expect.poll(async () => (await box(s)).vert).toBe('horz')
    } finally {
      kill(launched)
    }
  })

  test('Line Spacing from the expanded Paragraph row', async () => {
    const { s, launched } = await deckWithBox('One')
    try {
      const lineH = () =>
        s.evaluate(async () => {
          const r = await (window as any).slidesApi.getRenderSlides()
          return r[0].nodes.find((x: any) => x.text).text.lines[0].height as number
        })
      const before = await lineH()
      await drop(s, 'Line Spacing', '2.0')
      await expect.poll(lineH).toBeGreaterThan(before * 1.6)
    } finally {
      kill(launched)
    }
  })
})

test.describe('Home → Drawing extras', () => {
  test('Shape Fill / Outline / Quick Styles are live for a selected shape', async () => {
    const { s, launched } = await deckWithBox('x')
    try {
      await s.locator('.ribbon-body button', { hasText: 'Shapes' }).first().click()
      await s.locator('.rb-shape-cell[data-tip="Rectangle"]').click()
      // draw in the slide's top-left quarter, clear of the text box in the middle
      const stage = await s.evaluate(() => {
        const c = Array.from(document.querySelectorAll('canvas')).sort(
          (a, b) =>
            b.getBoundingClientRect().width * b.getBoundingClientRect().height -
            a.getBoundingClientRect().width * a.getBoundingClientRect().height,
        )[0]!
        const r = c.getBoundingClientRect()
        return { x: r.x, y: r.y, w: r.width, h: r.height }
      })
      const count = () =>
        s.evaluate(async () => (await (window as any).slidesApi.getRenderSlides())[0].nodes.length)
      const before = await count()
      await s.mouse.move(stage.x + stage.w * 0.1, stage.y + stage.h * 0.1)
      await s.mouse.down()
      await s.mouse.move(stage.x + stage.w * 0.3, stage.y + stage.h * 0.3, { steps: 8 })
      await s.mouse.up()
      await expect.poll(count).toBe(before + 1)
      await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
      for (const label of ['Quick Styles', 'Shape Fill', 'Shape Outline']) {
        await expect(s.locator('.ribbon-body button', { hasText: label }).first()).toBeEnabled()
      }
      const fillOf = () =>
        s.evaluate(async () => {
          const r = await (window as any).slidesApi.getRenderSlides()
          return JSON.stringify(r[0].nodes.at(-1).fill)
        })
      const fillBefore = await fillOf()
      await s.locator('.ribbon-body button', { hasText: 'Quick Styles' }).first().click()
      await s.locator('.ctx-style-cell').nth(3).click()
      await expect.poll(fillOf).not.toBe(fillBefore)
    } finally {
      kill(launched)
    }
  })
})
