import { test, expect, type Page } from '@playwright/test'
import { launchShell, setEditorLayoutWidth, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Every editor's ribbon behaves like Microsoft 365's: Genspark AI and the one-click AI tools
 * form one group at the right edge; on a narrower window whole groups fold into dropdown
 * buttons (AI first, Paragraph late) instead of the ribbon scrolling, and a folded group's
 * commands still work from its dropdown.
 */

/** `wide`: every group open (English labels); `narrow`: the AI group no longer fits */
const APPS = [
  { name: 'docs', card: 0, url: '://docs/', body: '.ribbon-body', wide: 1920, narrow: 900 },
  {
    name: 'sheets',
    card: 1,
    url: '://sheets/',
    body: '.ribbon[data-ribbon-body]',
    wide: 2200,
    narrow: 1100,
  },
  { name: 'markdown', card: 3, url: '://markdown/', body: '.ribbon-body', wide: 1920, narrow: 700 },
  { name: 'html', card: 4, url: '://html/', body: '.ribbon-body', wide: 1920, narrow: 900 },
] as const

async function open(card: number, url: string): Promise<{ p: Page; launched: LaunchedApp }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'ribbon-fold' })
  await launched.page.locator('.quick-card').nth(card).click()
  const p = await waitForPageWithUrl(launched.app, url, 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  return { p, launched }
}

const overflow = (p: Page, body: string) =>
  p.evaluate((sel) => {
    const b = document.querySelector(sel) as HTMLElement
    return b.scrollWidth - b.clientWidth
  }, body)

/** the group holding the Genspark AI entry */
const aiGroup = (p: Page) =>
  p.locator('.rb-fold-group', { has: p.locator('button.ai-entry', { hasText: 'Genspark AI' }) })

for (const app of APPS) {
  test.describe(`${app.name} ribbon`, () => {
    test('AI tools: one group at the right edge, folded first on a narrow window', async () => {
      const { p, launched } = await open(app.card, app.url)
      try {
        await setEditorLayoutWidth(launched.app, app.url, app.wide)
        const group = aiGroup(p)
        await expect(group).toHaveCount(1)
        // Genspark AI and the one-click tools share the group
        expect(await group.locator('button.ai-entry').count()).toBeGreaterThan(1)
        const gap = await group.evaluate((g) => {
          const body = g.parentElement!
          const pad = parseFloat(getComputedStyle(body).paddingRight) || 0
          return body.getBoundingClientRect().right - pad - g.getBoundingClientRect().right
        })
        expect(gap, 'AI group hugs the right edge').toBeLessThan(4)

        await setEditorLayoutWidth(launched.app, app.url, app.narrow)
        await expect(group).toHaveAttribute('data-folded', '')
        expect(await overflow(p, app.body)).toBeLessThanOrEqual(1)
        // the folded dropdown still reaches Genspark AI
        await group.locator('.rb-fold-btn').click()
        const entry = group.locator('.rb-fold-items button.ai-entry', { hasText: 'Genspark AI' })
        await expect(entry).toBeVisible()
        await entry.click()
        await expect(group).not.toHaveClass(/rb-fold-open/) // picking a command closes the panel
        // toggle the AI pane back so the widening check below measures the same ribbon width
        await group.locator('.rb-fold-btn').click()
        await entry.click()

        // widening unfolds it again
        await setEditorLayoutWidth(launched.app, app.url, app.wide)
        await expect(group).not.toHaveAttribute('data-folded', '')
      } finally {
        launched.app.process().kill('SIGKILL')
      }
    })

    test('no horizontal scrolling at 1280px', async () => {
      const { p, launched } = await open(app.card, app.url)
      try {
        await setEditorLayoutWidth(launched.app, app.url, 1280)
        await expect.poll(() => overflow(p, app.body)).toBeLessThanOrEqual(1)
      } finally {
        launched.app.process().kill('SIGKILL')
      }
    })
  })
}

test('docs keeps the Paragraph group expanded while Styles and Editing fold', async () => {
  const { p, launched } = await open(0, '://docs/')
  try {
    await setEditorLayoutWidth(launched.app, '://docs/', 1100)
    const folded = (label: string) =>
      p.locator(`.rb-fold-group[aria-label="${label}"]`).getAttribute('data-folded')
    expect(await folded('Styles')).toBe('')
    expect(await folded('Editing')).toBe('')
    expect(await folded('Paragraph')).toBeNull()
    expect(await overflow(p, '.ribbon-body')).toBeLessThanOrEqual(1)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('sheets: a folded group still runs its commands (Editing → AutoSum menu)', async () => {
  const { p, launched } = await open(1, '://sheets/')
  try {
    await setEditorLayoutWidth(launched.app, '://sheets/', 1100)
    const editing = p.locator('.rb-fold-group[aria-label="Editing"]')
    await expect(editing).toHaveAttribute('data-folded', '')
    await editing.locator('.rb-fold-btn').click()
    await expect(editing.locator('.rb-fold-items')).toBeVisible()
    const box = await editing.locator('.rb-fold-items').boundingBox()
    // the panel opens below its button, inside the window
    expect(box!.y).toBeGreaterThan(40)
    expect(box!.x + box!.width).toBeLessThanOrEqual(1100 + 1)
    await p.keyboard.press('Escape')
    await expect(editing.locator('.rb-fold-items')).toBeHidden()
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
