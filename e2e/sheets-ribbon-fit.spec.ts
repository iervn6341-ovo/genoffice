import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Sheets ribbon never scrolls sideways (Microsoft 365's behaviour): groups fold into dropdown
 * buttons, lowest priority first, on every tab — not only Home — and a folded group still runs its
 * commands.
 */
const TAB = 'nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)'
const BODY = '[data-ribbon-body]'

async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-ribbon-fit' })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector(BODY, { timeout: 20_000 })
  await p.waitForTimeout(1500)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')

for (const width of [1280, 1000, 900]) {
  test(`every Sheets tab fits at ${width}px without scrolling sideways`, async () => {
    const { launched, p } = await openSheets()
    try {
      await setEditorLayoutWidth(launched.app, '://sheets/', width)
      const tabs = (await p.locator(TAB).allInnerTexts()).map((s) => s.trim())
      expect(tabs.length).toBeGreaterThanOrEqual(7)
      const over: string[] = []
      for (const tab of tabs) {
        await p
          .locator(TAB, { hasText: new RegExp(`^${tab}$`) })
          .first()
          .click()
        await p.waitForTimeout(400)
        const m = await p.evaluate((sel) => {
          const el = document.querySelector(sel) as HTMLElement
          return { sw: el.scrollWidth, cw: el.clientWidth }
        }, BODY)
        if (m.sw > m.cw + 1) over.push(`${tab}: ${m.sw} > ${m.cw}`)
      }
      expect(over, `tabs that scroll sideways at ${width}px`).toEqual([])
    } finally {
      await kill(launched)
    }
  })
}

test('a folded Insert group opens as a flyout and its command still works', async () => {
  const { launched, p } = await openSheets()
  try {
    await setEditorLayoutWidth(launched.app, '://sheets/', 900)
    await p
      .locator(TAB, { hasText: /^Insert$/ })
      .first()
      .click()
    await p.waitForTimeout(400)
    const symbols = p
      .locator('.rb-fold-group[data-folded]')
      .filter({ hasText: /Symbol|Equation/ })
      .first()
    await expect(symbols).toBeVisible()
    await symbols.locator('.rb-fold-btn').click()
    await expect(symbols.locator('.rb-fold-items')).toBeVisible()
    await symbols
      .locator('.rb-fold-items button', { hasText: /Equation/ })
      .first()
      .click()
    await expect(p.locator('.dialog-backdrop')).toBeVisible()
  } finally {
    await kill(launched)
  }
})

test('nothing folds on a wide window', async () => {
  const { launched, p } = await openSheets()
  try {
    await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
    for (const tab of ['Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View']) {
      await p
        .locator(TAB, { hasText: new RegExp(`^${tab}$`) })
        .first()
        .click()
      await p.waitForTimeout(350)
      expect(await p.locator('.rb-fold-btn:visible').count(), tab).toBe(0)
    }
  } finally {
    await kill(launched)
  }
})

test('a ribbon dropdown reaching over the AI panel edge is clickable all across (the resize handle stays below it)', async () => {
  const { launched, p } = await openSheets()
  try {
    await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
    await p
      .locator(TAB, { hasText: /^Page Layout$/ })
      .first()
      .click()
    await p.waitForTimeout(400)
    await p.locator(`${BODY} [aria-label="Size"]`).first().click()
    await p.waitForTimeout(300)
    // the options whose middle is where the panel's 6px resize handle is (the layout puts the
    // dropdown across the panel's right edge) are still the top element there
    const options = p.locator('.menu-select-drop button')
    const count = await options.count()
    expect(count).toBeGreaterThanOrEqual(7)
    const covered = await p.evaluate(() => {
      const handle = document.querySelector('.ai-panel-resizer')?.getBoundingClientRect()
      if (!handle) return ['no resize handle']
      const bad: string[] = []
      for (const button of document.querySelectorAll('.menu-select-drop button')) {
        const box = button.getBoundingClientRect()
        // every pixel column of the handle that the option spans
        for (const x of [handle.x + 1, handle.x + handle.width / 2, handle.x + handle.width - 1]) {
          if (x < box.x || x > box.x + box.width) continue
          const top = document.elementFromPoint(x, box.y + box.height / 2)
          if (top !== button && !button.contains(top)) bad.push(`${button.textContent}@${x}`)
        }
      }
      return bad
    })
    expect(covered).toEqual([])
  } finally {
    await kill(launched)
  }
})
