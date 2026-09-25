import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Slides Home tab follows Microsoft PowerPoint's layout (dark Mac build used as the reference):
 * Clipboard → Slides → Font → Paragraph → Drawing (Pictures, Shapes, Text Box, Arrange) → …, with
 * the AI group where Microsoft 365 puts Copilot, at the far right. "From Current Slide" lives on
 * the Slide Show tab, not on Home.
 */

async function openSlides(width = 1900): Promise<{ s: Page; launched: LaunchedApp }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-ribbon' })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await s.waitForSelector('.ribbon', { timeout: 15_000 })
  await setEditorLayoutWidth(launched.app, '://slides/', width)
  await s.waitForTimeout(600)
  return { s, launched }
}

/** left edge (CSS px) of the first visible Home-ribbon control whose text/tip matches */
const leftOf = (s: Page, text: RegExp) =>
  s.evaluate((src) => {
    const re = new RegExp(src)
    for (const el of Array.from(document.querySelectorAll('.ribbon-body button'))) {
      const label = `${(el as HTMLElement).dataset.tip ?? ''} ${el.textContent ?? ''}`
      if (re.test(label)) {
        const r = el.getBoundingClientRect()
        if (r.width > 0) return r.left
      }
    }
    return null
  }, text.source)

test.describe('slides Home ribbon layout (PowerPoint order)', () => {
  test('groups appear in PowerPoint order with AI at the far right', async () => {
    const { s, launched } = await openSlides()
    try {
      const order = [
        '^\\s*Paste',
        'New Slide',
        'Pictures',
        'Shapes',
        'Text Box',
        'Arrange',
        'Find',
        'Genspark AI',
      ]
      const xs: number[] = []
      for (const label of order) {
        const x = await leftOf(s, new RegExp(label))
        expect(x, `"${label}" is on the Home tab`).not.toBeNull()
        xs.push(x!)
      }
      expect(xs, 'left-to-right order').toEqual([...xs].sort((a, b) => a - b))
      // the AI group hugs the right edge of the ribbon
      const gap = await s.evaluate(() => {
        const ai = Array.from(document.querySelectorAll('.ribbon-body button')).find((b) =>
          /Genspark AI/.test(b.textContent ?? ''),
        )!
        const group = ai.closest('.ribbon-group') as HTMLElement
        const body = document.querySelector('.ribbon-body') as HTMLElement
        return body.getBoundingClientRect().right - group.getBoundingClientRect().left
      })
      expect(gap).toBeLessThan(500) // 4 AI buttons wide, not stranded mid-ribbon
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('"From Current Slide" moved off Home to the Slide Show tab', async () => {
    const { s, launched } = await openSlides()
    try {
      expect(await leftOf(s, /From Current Slide/)).toBeNull()
      await s.locator('.ribbon-tab', { hasText: 'Slide Show' }).click()
      await s.waitForTimeout(300)
      expect(await leftOf(s, /From Current Slide/)).not.toBeNull()
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('the Slides group has Layout / Reset / Add Section rows', async () => {
    const { s, launched } = await openSlides()
    try {
      const rows = s.locator('.rb-slides-col button')
      await expect(rows).toHaveCount(3)
      await expect(rows.nth(0)).toContainText('Layout')
      await expect(rows.nth(1)).toContainText('Reset')
      await expect(rows.nth(2)).toContainText('Section')
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('Text Box from Home inserts a box and starts typing in it', async () => {
    const { s, launched } = await openSlides()
    try {
      await s.locator('.ribbon-body button', { hasText: 'Text Box' }).first().click()
      await s.waitForSelector('.slide-text-editor', { timeout: 5000 })
      await s.keyboard.type('Hi')
      expect(
        await s.evaluate(() => document.querySelector('.slide-text-editor')?.textContent),
      ).toBe('Hi')
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('Shapes from Home opens the gallery and draws the chosen shape', async () => {
    const { s, launched } = await openSlides()
    try {
      const count = () =>
        s.evaluate(async () => (await (window as any).slidesApi.getRenderSlides())[0].nodes.length)
      const before = await count()
      await s.locator('.ribbon-body button', { hasText: 'Shapes' }).first().click()
      await s.locator('.rb-shape-cell[data-tip="Rectangle"]').click()
      await s.mouse.move(640, 300)
      await s.mouse.down()
      await s.mouse.move(900, 400, { steps: 8 })
      await s.mouse.up()
      await s.waitForTimeout(600)
      expect(await count()).toBe(before + 1)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })
})

test.describe('narrow windows fold groups instead of scrolling', () => {
  const overflow = (s: Page) =>
    s.evaluate(() => {
      const b = document.querySelector('.ribbon-body') as HTMLElement
      return b.scrollWidth - b.clientWidth
    })

  test('at 1200px the Home tab fits without a scrollbar and the folded Drawing group still works', async () => {
    const { s, launched } = await openSlides(1200)
    try {
      await expect.poll(() => overflow(s)).toBeLessThanOrEqual(1)
      const drawing = s.locator('[data-rbgroup="drawing"] button.rb-big').first()
      await expect(drawing).toContainText('Drawing')
      const count = () =>
        s.evaluate(async () => (await (window as any).slidesApi.getRenderSlides())[0].nodes.length)
      const before = await count()
      await drawing.click()
      await s.locator('.rb-collapse-panel button', { hasText: 'Shapes' }).first().click()
      await s.locator('.rb-shape-cell[data-tip="Rectangle"]').click()
      const stage = await s.evaluate(() => {
        const c = Array.from(document.querySelectorAll('canvas')).sort(
          (a, b) =>
            b.getBoundingClientRect().width * b.getBoundingClientRect().height -
            a.getBoundingClientRect().width * a.getBoundingClientRect().height,
        )[0]!
        const r = c.getBoundingClientRect()
        return { x: r.x, y: r.y, w: r.width, h: r.height }
      })
      await s.mouse.move(stage.x + stage.w * 0.2, stage.y + stage.h * 0.2)
      await s.mouse.down()
      await s.mouse.move(stage.x + stage.w * 0.4, stage.y + stage.h * 0.4, { steps: 8 })
      await s.mouse.up()
      await expect.poll(count).toBe(before + 1)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('only after everything is folded does the ribbon scroll', async () => {
    const { s, launched } = await openSlides(800)
    try {
      // at 800px even the fully folded Home tab is too wide: every fold group must be folded
      if ((await overflow(s)) > 1) {
        for (const g of ['paragraph', 'drawing', 'aiTools', 'slides']) {
          await expect(
            s.locator(`[data-rbgroup="${g}"] .ribbon-group-items > .rb-drop-wrap`).first(),
          ).toBeVisible()
        }
      }
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('Paragraph stays expanded after AI Tools and Drawing fold (1280px)', async () => {
    const { s, launched } = await openSlides(1280)
    try {
      await expect(s.locator('[data-rbgroup="paragraph"] .rb-para-inline')).toHaveCount(1)
      await expect(s.locator('[data-rbgroup="aiTools"] .rb-drop-wrap')).toHaveCount(1)
      expect(await overflow(s)).toBeLessThanOrEqual(1)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })

  test('Genspark AI sits inside the AI Tools group', async () => {
    const { s, launched } = await openSlides(2200)
    try {
      const group = s.locator('[data-rbgroup="aiTools"]')
      await expect(group.locator('button.ai-entry', { hasText: 'Genspark AI' })).toHaveCount(1)
      expect(await group.locator('button.ai-entry').count()).toBe(4)
    } finally {
      launched.app.process().kill('SIGKILL')
    }
  })
})

test('widening the window unfolds the groups again', async () => {
  const { s, launched } = await openSlides(1000)
  try {
    await expect(s.locator('[data-rbgroup="paragraph"] .rb-para-inline')).toHaveCount(0)
    await setEditorLayoutWidth(launched.app, '://slides/', 2200)
    await expect(s.locator('[data-rbgroup="paragraph"] .rb-para-inline')).toHaveCount(1)
    await expect(s.locator('.rb-slides-col')).toHaveCount(1)
    await expect(s.locator('.ribbon-body button', { hasText: 'Shape Fill' }).first()).toBeVisible()
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
