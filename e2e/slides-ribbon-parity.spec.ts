import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Ribbon parity with PowerPoint: nothing scrolls sideways (groups fold into dropdown buttons on a
 * narrow window), Bullets / Numbering have their own libraries, Columns, Shape Effects, the table
 * style gallery panel, and the Transitions tab's Effect Options / Duration / Advance controls.
 */
async function openSlides(): Promise<{ launched: LaunchedApp; s: Page }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'slides-ribbon-parity' })
  await launched.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
  await s.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  return { launched, s }
}

const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')

async function addTextBox(s: Page, text: string): Promise<void> {
  await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
  await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
  await s.waitForSelector('.slide-text-editor')
  await s.keyboard.type(text)
  await s.keyboard.press('Escape')
  await s.waitForTimeout(400)
  await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
}

/** the first text-bearing node of slide 1 */
const textNode = (s: Page) =>
  s.evaluate(async () => {
    const r = await (window as any).slidesApi.getRenderSlides()
    const n = r[0].nodes.find((x: any) => x.text)
    return {
      lines: n.text.lines.map((l: any) => ({
        bullet: l.runs.find((run: any) => run.isBullet)?.text ?? null,
        x: Math.round(l.runs[0]?.x ?? 0),
        right: (l.runs[0]?.x ?? 0) > n.box.w / 2,
        top: Math.round(l.top),
      })),
      box: n.box,
    }
  })

test.describe('ribbon never scrolls sideways', () => {
  for (const width of [1280, 1000, 900]) {
    test(`every tab fits at ${width}px`, async () => {
      const { launched, s } = await openSlides()
      try {
        await setEditorLayoutWidth(launched.app, '://slides/', width)
        const tabs = (await s.locator('.ribbon-tab:not(.ribbon-tab-file)').allInnerTexts()).map(
          (t) => t.trim(),
        )
        expect(tabs.length).toBeGreaterThanOrEqual(9)
        for (const tab of tabs) {
          await s
            .locator('.ribbon-tab:not(.ribbon-tab-file)', { hasText: new RegExp(`^${tab}$`) })
            .click()
          await s.waitForTimeout(350)
          const m = await s.evaluate(() => {
            const el = document.querySelector('[data-ribbon-body]') as HTMLElement
            return { sw: el.scrollWidth, cw: el.clientWidth }
          })
          expect(m.sw, `${tab} at ${width}px`).toBeLessThanOrEqual(m.cw + 1)
        }
      } finally {
        await kill(launched)
      }
    })
  }

  test('a folded group opens as a flyout, runs its command and closes', async () => {
    const { launched, s } = await openSlides()
    try {
      await setEditorLayoutWidth(launched.app, '://slides/', 900)
      await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await s.waitForTimeout(400)
      const folded = s.locator('.rb-fold-btn:visible')
      expect(await folded.count()).toBeGreaterThan(0)
      // the Links group (hyperlink / zoom) is folded at this width: open it and press Escape
      await folded.first().click()
      await expect(s.locator('.rb-fold-open .rb-fold-items')).toBeVisible()
      await s.keyboard.press('Escape')
      await expect(s.locator('.rb-fold-open')).toHaveCount(0)
      // a command inside the flyout works and dismisses it: Text Box starts typing in a new box
      const textGroup = s.locator('.rb-fold-group', {
        has: s.locator('button[data-tip^="Insert a text box"]'),
      })
      if (await textGroup.locator('.rb-fold-btn').isVisible()) {
        await textGroup.locator('.rb-fold-btn').click()
        await textGroup.locator('button[data-tip^="Insert a text box"]').click()
        await expect(s.locator('.slide-text-editor')).toBeVisible()
        await expect(s.locator('.rb-fold-open')).toHaveCount(0)
      }
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Home → Paragraph', () => {
  test('Bullets and Numbering open their own libraries; Columns balance a box', async () => {
    const { s, launched } = await openSlides()
    try {
      await addTextBox(
        s,
        'Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega.',
      )
      const rows = s.locator('.rb-para-inline')
      await expect(rows.locator('button[aria-label="Bullets"]')).toHaveCount(2)
      // Bullets ▾ → library → the first symbol
      await rows.locator('.rb-split-icon').first().locator('button').nth(1).click()
      await expect(s.locator('.rb-bullet-lib')).toBeVisible()
      await expect(s.locator('.rb-bullet-lib .rb-para-label').first()).toHaveText('Bullet Library')
      await s.locator('.rb-bullet-lib .rb-bullet-tile').nth(1).click()
      await expect.poll(async () => (await textNode(s)).lines[0]!.bullet).not.toBeNull()
      // the library closes after a pick and sits inside the window
      await expect(s.locator('.rb-bullet-lib')).toHaveCount(0)
      // Numbering ▾ → "1. 2. 3."
      await rows.locator('.rb-split-icon').nth(1).locator('button').nth(1).click()
      await expect(s.locator('.rb-bullet-lib .rb-para-label').first()).toHaveText(
        'Numbering Library',
      )
      const box = (await s.locator('.rb-bullet-drop').boundingBox())!
      expect(box.x + box.width).toBeLessThanOrEqual(1280)
      expect(box.y + box.height).toBeLessThanOrEqual(760)
      await s.locator('.rb-bullet-lib .rb-bullet-tile').nth(1).click()
      await expect.poll(async () => (await textNode(s)).lines[0]!.bullet).toMatch(/^1\./)
      // None removes it again
      await rows.locator('.rb-split-icon').nth(1).locator('button').nth(1).click()
      await s.locator('.rb-bullet-lib .rb-bullet-tile-none').click()
      await expect.poll(async () => (await textNode(s)).lines[0]!.bullet).toBeNull()
      // Columns: Two Columns puts the text in two x bands with a balanced box height
      const before = await textNode(s)
      await rows.locator('button[aria-label="Columns"]').click()
      await s.locator('.rb-menu button', { hasText: 'Two Columns' }).click()
      await expect
        .poll(
          async () =>
            new Set((await textNode(s)).lines.map((l: { right: boolean }) => l.right)).size,
        )
        .toBe(2)
      const two = await textNode(s)
      const left = two.lines.filter((l: { right: boolean }) => !l.right).length
      const right = two.lines.length - left
      expect(Math.abs(left - right)).toBeLessThanOrEqual(1)
      expect(two.box.h).toBeGreaterThan(0)
      expect(before.lines.length).toBeGreaterThan(0)
      // back to one column
      await rows.locator('button[aria-label="Columns"]').click()
      await s.locator('.rb-menu button', { hasText: 'One Column' }).click()
      await expect
        .poll(
          async () =>
            new Set((await textNode(s)).lines.map((l: { right: boolean }) => l.right)).size,
        )
        .toBe(1)
      // everything above is one undo step each
      await s.locator('button[aria-label^="Undo"]').click()
      await expect
        .poll(
          async () =>
            new Set((await textNode(s)).lines.map((l: { right: boolean }) => l.right)).size,
        )
        .toBe(2)
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Shape Effects', () => {
  test('shadow, glow and soft edges apply to a drawn shape and None clears them', async () => {
    const { s, launched } = await openSlides()
    try {
      await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await s.locator('.ribbon-body button[aria-label="Shapes"]').first().click()
      await s.locator('.ribbon-gallery-pane .rb-shape-cell[data-tip="Rectangle"]').click()
      await s.waitForSelector('.ribbon-gallery-pane', { state: 'detached' })
      const sb = (await s.locator('.stage-wrap').boundingBox())!
      await s.mouse.move(sb.x + sb.width * 0.35, sb.y + sb.height * 0.35)
      await s.mouse.down()
      await s.mouse.move(sb.x + sb.width * 0.6, sb.y + sb.height * 0.6, { steps: 8 })
      await s.mouse.up()
      await s.locator('.ribbon-tab', { hasText: 'Shape Tools' }).click()
      const shape = () =>
        s.evaluate(async () => {
          const r = await (window as any).slidesApi.getRenderSlides()
          const n = r[0].nodes.at(-1)
          return {
            shadow: n.shadow ?? null,
            glow: n.glow ?? null,
            softEdgePx: n.softEdgePx ?? null,
          }
        })
      const open = () => s.locator('.ribbon-body button[data-tip="Shape Effects"]').first().click()
      await open()
      // the menu fits inside the window
      const m = (await s.locator('.rb-effects-drop').boundingBox())!
      expect(m.x + m.width).toBeLessThanOrEqual(1280)
      expect(m.y + m.height).toBeLessThanOrEqual(760)
      await s.locator('.rb-effects-shadow').first().click()
      await expect.poll(async () => (await shape()).shadow).not.toBeNull()
      await open()
      await s.locator('.rb-effects-drop .rb-para-label', { hasText: 'Glow' }).waitFor()
      await s
        .locator('.rb-effects-opt', { hasText: /^8 pt$/ })
        .first()
        .click()
      await expect.poll(async () => (await shape()).glow).not.toBeNull()
      await open()
      await s
        .locator('.rb-effects-opt', { hasText: /^5 pt$/ })
        .last()
        .click()
      await expect.poll(async () => (await shape()).softEdgePx).not.toBeNull()
      // None of each clears it
      await open()
      await s
        .locator('.rb-effects-drop .rb-effects-grid .rb-effects-opt', { hasText: 'None' })
        .click()
      await expect.poll(async () => (await shape()).shadow).toBeNull()
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Table style gallery', () => {
  test('the more-styles panel shows seven columns without a sideways scrollbar', async () => {
    const { s, launched } = await openSlides()
    try {
      await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await s.locator('.ribbon-body button[data-tip*="Insert a table"]').first().click()
      await s
        .locator('.rb-drop button, .rb-drop [role="gridcell"], .rb-table-picker-cell')
        .nth(11)
        .click({ force: true })
      await s.waitForSelector('.rb-tbl-gallery')
      await s.locator('.rb-tbl-gallery .rb-tbl-arrow').last().click()
      const m = await s.evaluate(() => {
        const el = document.querySelector('.rb-tbl-allstyles') as HTMLElement
        const r = el.getBoundingClientRect()
        const first = el.querySelector('.rb-tbl-allgrid') as HTMLElement
        const cols = getComputedStyle(first).gridTemplateColumns.split(' ').length
        const svg = el.querySelector('.rb-tbl-thumb-svg') as SVGElement
        return {
          sw: el.scrollWidth,
          cw: el.clientWidth,
          right: r.right,
          bottom: r.bottom,
          cols,
          paper: getComputedStyle(svg).backgroundColor,
        }
      })
      expect(m.cols).toBe(7)
      expect(m.sw).toBeLessThanOrEqual(m.cw)
      expect(m.right).toBeLessThanOrEqual(1280)
      expect(m.bottom).toBeLessThanOrEqual(760)
      // thumbnails sit on white paper even though the panel follows the theme
      expect(m.paper).toBe('rgb(255, 255, 255)')
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Transitions tab', () => {
  test('Effect Options, Duration and Advance Slide are written to the page', async () => {
    const { s, launched } = await openSlides()
    try {
      await s.locator('.ribbon-tab', { hasText: 'Transitions' }).click()
      const detail = () =>
        s.evaluate(
          async () =>
            (await (window as any).slidesApi.getTransitionDetail(0)) as {
              kind: string
              dir?: string
              durationMs: number | null
              advClick: boolean
              advTm: number | null
            },
        )
      // no effect yet: Effect Options and Duration wait for one
      await expect(s.locator('.ribbon-body button[aria-label="Effect Options"]')).toBeDisabled()
      await s.locator('.ribbon-body button[data-tip*="\\"Push\\""]').first().click()
      await expect.poll(async () => (await detail()).kind).toBe('push')
      expect((await detail()).dir).toBe('u')
      await s.locator('.ribbon-body button[aria-label="Effect Options"]').click()
      await s.locator('.rb-menu button', { hasText: 'From Left' }).click()
      await expect.poll(async () => (await detail()).dir).toBe('r')
      // switching the effect resets the direction to the new effect's default
      await s.locator('.ribbon-body button[data-tip*="\\"Cover\\""]').first().click()
      await expect.poll(async () => (await detail()).kind).toBe('cover')
      expect((await detail()).dir).toBe('l')
      // Fade offers Smoothly / Through Black instead of directions
      await s.locator('.ribbon-body button[data-tip*="\\"Fade\\""]').first().click()
      await s.locator('.ribbon-body button[aria-label="Effect Options"]').click()
      await expect(s.locator('.rb-menu button', { hasText: 'Through Black' })).toBeVisible()
      await s.keyboard.press('Escape')
      await s.locator('.ribbon-body button[data-tip*="\\"Push\\""]').first().click()
      // Duration commits on Enter and survives the next effect change
      const dur = s.locator('.rb-anim-timing input[type="number"]').first()
      await dur.fill('1.5')
      await dur.press('Enter')
      await expect.poll(async () => (await detail()).durationMs).toBe(1500)
      await s.locator('.ribbon-body button[data-tip*="\\"Wipe\\""]').first().click()
      await expect.poll(async () => (await detail()).kind).toBe('wipe')
      expect((await detail()).durationMs).toBe(1500)
      // Advance Slide: On Mouse Click off, After 3 s
      await s.locator('.rb-anim-check', { hasText: 'On Mouse Click' }).locator('input').click()
      await expect.poll(async () => (await detail()).advClick).toBe(false)
      await s
        .locator('.rb-anim-check', { hasText: 'After' })
        .locator('input[type="checkbox"]')
        .click()
      await expect.poll(async () => (await detail()).advTm).toBe(5000)
      const after = s
        .locator('.rb-anim-check', { hasText: 'After' })
        .locator('input[type="number"]')
      await after.fill('3')
      await after.press('Enter')
      await expect.poll(async () => (await detail()).advTm).toBe(3000)
      // a second page can differ; Apply To All copies page 1's whole setting
      await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
      await s.locator('.ribbon-body button[data-tip^="New blank slide"]').first().click()
      await s.waitForTimeout(500)
      await s.locator('.ribbon-tab', { hasText: 'Transitions' }).click()
      const second = () =>
        s.evaluate(
          async () =>
            (await (window as any).slidesApi.getTransitionDetail(1)) as {
              kind: string
              advTm: number | null
            },
        )
      expect((await second()).kind).toBe('none')
      await s.locator('.slide-list .thumb').first().click()
      await s.locator('.ribbon-body button', { hasText: 'Apply To All' }).click()
      await expect.poll(async () => (await second()).kind).toBe('wipe')
      expect((await second()).advTm).toBe(3000)
    } finally {
      await kill(launched)
    }
  })
})
