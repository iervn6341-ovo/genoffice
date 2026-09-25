import { test, expect, type Page } from '@playwright/test'
import { PNG } from 'pngjs'
import { launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Slides text editing: the DOM editor must sit exactly where the canvas draws the text, so the
 * caret never looks displaced and nothing jumps when editing starts or ends.
 *
 * Method: measure the dark ink of the text box with the editor open (DOM) and again after Esc
 * (canvas) — the two must coincide. The selection frame is canvas-drawn, so the measured clip is
 * the DOM text extent, which stays clear of the frame's handles for short left-aligned text.
 */

interface Rect {
  x: number
  y: number
  w: number
  h: number
}

async function inkBounds(p: Page, clip: Rect): Promise<Rect | null> {
  const buf = await p.screenshot({
    clip: { x: clip.x, y: clip.y, width: clip.w, height: clip.h },
    timeout: 8000,
  })
  const png = PNG.sync.read(buf)
  const sx = png.width / clip.w
  const sy = png.height / clip.h
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -1
  let y1 = -1
  for (let y = 0; y < png.height; y++) {
    for (let x = 0; x < png.width; x++) {
      const i = (y * png.width + x) * 4
      const lum = 0.299 * png.data[i]! + 0.587 * png.data[i + 1]! + 0.114 * png.data[i + 2]!
      if (lum < 110 && png.data[i + 3]! > 200) {
        x0 = Math.min(x0, x)
        x1 = Math.max(x1, x)
        y0 = Math.min(y0, y)
        y1 = Math.max(y1, y)
      }
    }
  }
  if (x1 < 0) return null
  return { x: clip.x + x0 / sx, y: clip.y + y0 / sy, w: (x1 - x0 + 1) / sx, h: (y1 - y0 + 1) / sy }
}

const editorRect = (s: Page): Promise<Rect> =>
  s.evaluate(() => {
    const r = document.querySelector('.slide-text-editor')!.getBoundingClientRect()
    return { x: r.x, y: r.y, w: r.width, h: r.height }
  })

/** union of the client rects of everything the editor lays out */
const textExtent = (s: Page): Promise<Rect> =>
  s.evaluate(() => {
    const range = document.createRange()
    range.selectNodeContents(document.querySelector('.slide-text-editor')!)
    const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0)
    const x0 = Math.min(...rects.map((r) => r.left))
    const y0 = Math.min(...rects.map((r) => r.top))
    return {
      x: x0,
      y: y0,
      w: Math.max(...rects.map((r) => r.right)) - x0,
      h: Math.max(...rects.map((r) => r.bottom)) - y0,
    }
  })

const tip = (s: Page, t: string) => s.locator(`.ribbon-body button[data-tip^="${t}"]`).first()

async function newDeckWithTextBox(launched: LaunchedApp, text: string): Promise<Page> {
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
  await tip(s, 'Insert a text box').click()
  await s.waitForSelector('.slide-text-editor', { timeout: 5000 })
  await s.keyboard.type(text)
  await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
  return s
}

/** ink with the editor open vs after Esc (canvas), both over the same DOM-derived clip */
async function domVsCanvas(s: Page) {
  const rect = await editorRect(s)
  const ext = await textExtent(s)
  const left = rect.x + 6
  const clip = { x: left, y: ext.y, w: ext.x + ext.w + 3 - left, h: ext.h }
  const dom = await inkBounds(s, clip)
  await s.keyboard.press('Escape')
  await s.waitForTimeout(500)
  const canvas = await inkBounds(s, clip)
  expect(dom, 'DOM ink').not.toBeNull()
  expect(canvas, 'canvas ink').not.toBeNull()
  return { dom: dom!, canvas: canvas!, box: rect }
}

/** the editor and the canvas may differ by sub-pixel rounding only */
const TOL = 1

async function withDeck<T>(text: string, fn: (s: Page, l: LaunchedApp) => Promise<T>): Promise<T> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-text-editing' })
  try {
    const s = await newDeckWithTextBox(launched, text)
    return await fn(s, launched)
  } finally {
    // closing with editor tabs open waits on a native dialog: end the process instead
    launched.app.process().kill('SIGKILL')
  }
}

test.describe('slides text editing: editor matches canvas', () => {
  test('plain text', async () => {
    await withDeck('Hi', async (s) => {
      const { dom, canvas } = await domVsCanvas(s)
      expect(Math.abs(dom.x - canvas.x)).toBeLessThanOrEqual(TOL)
      expect(Math.abs(dom.y - canvas.y)).toBeLessThanOrEqual(TOL)
      expect(Math.abs(dom.w - canvas.w)).toBeLessThanOrEqual(TOL)
    })
  })

  test('CJK and mixed text', async () => {
    await withDeck('Hi 你好 World', async (s) => {
      const { dom, canvas } = await domVsCanvas(s)
      expect(Math.abs(dom.x - canvas.x)).toBeLessThanOrEqual(TOL)
      expect(Math.abs(dom.y - canvas.y)).toBeLessThanOrEqual(TOL)
      expect(Math.abs(dom.w - canvas.w)).toBeLessThanOrEqual(TOL)
    })
  })

  for (const pct of ['75%', '100%', '150%']) {
    test(`bullet toggled while editing at ${pct}: text does not shift on commit`, async () => {
      await withDeck('Hi', async (s) => {
        await tip(s, 'Paragraph').click()
        await s.locator('.rb-bullet-tile').nth(1).click()
        await s.locator('.rb-bullet-hang', { hasText: pct }).click()
        await tip(s, 'Paragraph').click()
        await s.waitForTimeout(300)
        const { dom, canvas } = await domVsCanvas(s)
        // the whole run (bullet + text) must end where the canvas ends: a too-narrow preview
        // indent used to leave the text ~5px left of its committed position
        expect(Math.abs(dom.x - canvas.x)).toBeLessThanOrEqual(TOL)
        expect(Math.abs(dom.w - canvas.w)).toBeLessThanOrEqual(TOL)
        expect(Math.abs(dom.y - canvas.y)).toBeLessThanOrEqual(TOL)
      })
    })
  }

  test('bullet size chosen while editing survives the commit', async () => {
    await withDeck('Hi', async (s) => {
      await tip(s, 'Paragraph').click()
      await s.locator('.rb-bullet-tile').nth(1).click()
      await s.locator('.rb-bullet-hang', { hasText: '150%' }).click()
      await tip(s, 'Paragraph').click()
      const bulletSize = () =>
        s.evaluate(
          () =>
            document
              .querySelector('.slide-text-editor')
              ?.firstElementChild?.getAttribute('style')
              ?.match(/--bullet-size:\s*([\d.]+)px/)?.[1],
        )
      const before = Number(await bulletSize())
      const box = await editorRect(s)
      await s.keyboard.press('Escape')
      await s.waitForTimeout(500)
      await s.mouse.dblclick(box.x + 40, box.y + 12)
      await s.waitForSelector('.slide-text-editor', { timeout: 3000 })
      await s.waitForTimeout(300)
      const after = Number(await bulletSize())
      expect(before).toBeGreaterThan(24) // 150% of the 24px text
      expect(after).toBeCloseTo(before, 0)
    })
  })

  test('preview bullet dot stays centred on the text like the canvas dot', async () => {
    await withDeck('Hi', async (s) => {
      await tip(s, 'Paragraph').click()
      await s.locator('.rb-bullet-tile').nth(1).click()
      await s.locator('.rb-bullet-hang', { hasText: '150%' }).click()
      await tip(s, 'Paragraph').click()
      await s.waitForTimeout(300)
      const shift = await s.evaluate(() =>
        parseFloat(
          (
            document.querySelector('.slide-text-editor')?.firstElementChild as HTMLElement
          ).style.getPropertyValue('--bullet-shift'),
        ),
      )
      // 150% of 24px text → the glyph grows 12px; centre it 0.35em of that lower
      expect(shift).toBeCloseTo(12 * 0.35, 1)
    })
  })
})

/** the last node of slide 1 as the renderer sees it (geometry + first text line) */
const lastNode = (s: Page) =>
  s.evaluate(async () => {
    const r = await (window as any).slidesApi.getRenderSlides()
    const n = r[0].nodes.at(-1)
    const line = n.text?.lines?.[0]
    const run = line?.runs?.find((x: any) => !x.isBullet) ?? line?.runs?.[0]
    return {
      type: n.type as string,
      box: n.box as { x: number; y: number; w: number; h: number },
      anchor: n.text?.anchor as string | undefined,
      insetX: (n.text?.insets?.l ?? 0) + (n.text?.insets?.r ?? 0),
      align: line?.align as string | undefined,
      color: run?.color as string | undefined,
      run: run
        ? { x: run.x as number, w: run.widthPx as number, base: run.baselineY as number }
        : null,
    }
  })

test.describe('slides: PowerPoint editing conventions', () => {
  test('Enter and F2 start editing the selected shape with the caret at the end', async () => {
    await withDeck('abc', async (s) => {
      await s.keyboard.press('Escape') // leave editing; the box stays selected
      await s.waitForTimeout(400)
      expect(await s.evaluate(() => !!document.querySelector('.slide-text-editor'))).toBe(false)
      for (const key of ['Enter', 'F2']) {
        await s.keyboard.press(key)
        await s.waitForSelector('.slide-text-editor', { timeout: 3000 })
        // typing continues after the existing text
        await s.keyboard.type('X')
        expect(
          await s.evaluate(() => document.querySelector('.slide-text-editor')?.textContent),
        ).toMatch(/abcX$/)
        await s.keyboard.press('Escape')
        await s.waitForTimeout(400)
        // remove the X again so both keys start from the same text
        await s.keyboard.press('Enter')
        await s.waitForSelector('.slide-text-editor', { timeout: 3000 })
        await s.keyboard.press('Backspace')
        await s.keyboard.press('Escape')
        await s.waitForTimeout(400)
      }
    })
  })

  test('a new rectangle holds white, centred, vertically middle text', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-text-editing' })
    try {
      await launched.page.locator('.quick-card').nth(2).click()
      const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
      await s.waitForSelector('.ribbon', { timeout: 15_000 })
      await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await tip(s, 'Insert a shape').click()
      await s.locator('.rb-shape-cell[data-tip="Rectangle"]').click()
      await s.mouse.move(640, 300)
      await s.mouse.down()
      await s.mouse.move(900, 400, { steps: 8 })
      await s.mouse.up()
      await s.waitForTimeout(500)
      await s.mouse.dblclick(770, 350)
      await s.waitForSelector('.slide-text-editor', { timeout: 5000 })
      await s.keyboard.type('Hi')
      await s.keyboard.press('Escape')
      await s.waitForTimeout(600)
      const n = await lastNode(s)
      expect(n.type).toBe('shape')
      expect(n.align).toBe('center')
      expect(n.anchor).toBe('middle')
      expect(n.color?.toUpperCase()).toBe('#FFFFFF')
      // centred: run x is relative to the text area, i.e. the shape minus its left/right insets
      expect(Math.abs(n.run!.x + n.run!.w / 2 - (n.box.w - n.insetX) / 2)).toBeLessThan(3)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })
})
